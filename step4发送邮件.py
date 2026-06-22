import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
import mimetypes
import datetime
import pandas as pd
# ============= 一、需要按实际情况修改的配置 =============


now = datetime.datetime.now()
today =datetime.datetime(now.year,now.month,now.day).date().strftime('%Y-%m-%d')

data_日期 =pd.read_excel('交易日历_更新.xlsx')
DATE_STR=data_日期['上一个交易日'][0].strftime('%Y-%m-%d') 

# 图和 Excel 所在文件夹
FOLDER = Path(f"./邮件图片文件夹/{DATE_STR}")  # 请改为实际路径

# 发件邮箱相关（下面是示例，请按你自己的邮箱改）
SMTP_HOST = "smtp.qq.com"   # 如：smtp.qq.com / smtp.163.com / smtp.office365.com
SMTP_PORT = 465                   # SSL 一般 465，STARTTLS 一般 587
USE_SSL = True                    # 465 通常用 True；如果用 587 + STARTTLS，改为 False

USERNAME = ""   # 登录账号，一般就是邮箱地址
PASSWORD = "      # 很多邮箱是“客户端授权码”

FROM_ADDR = "2489592676@qq.com"
TO_ADDRS =["","",""]# # 收件人列表
CC_ADDRS = []  # 如有抄送，在这里加



# 业务相关信息
#DATE_STR = "2025-12-2
SUBJECT = f"{DATE_STR} 科创债ETF成分券估值相关信息"

GREETING = "各位领导、同事早上好："
SIGNATURE = ""#"—— XXX\n联系方式：XXXX-XXXXXXX"

# =====================================================


def collect_files(folder: Path):
    """收集 png 图和 Excel 文件，并做简单分类"""
    pngs = sorted(folder.glob("*.png"))
    excels = sorted(folder.glob("*.xls*"))

    # 包含“中短票据”的图片作为第1部分，其余为第2部分
    top_imgs = [p for p in pngs if "中短票据" in p.name]
    term_imgs = [p for p in pngs if "中短票据" not in p.name]

    # 兜底：如果没识别到“中短票据”，就前两张放第1部分
    if not top_imgs and len(pngs) >= 2:
        top_imgs = pngs[:2]
        term_imgs = pngs[2:]

    return pngs, top_imgs, term_imgs, excels


def build_table(img_paths, cid_map, cols=2, width=480):
    """按 N 列把图片排成 HTML 表格"""
    rows_html = []
    for i in range(0, len(img_paths), cols):
        tds = []
        for p in img_paths[i:i + cols]:
            cid = cid_map[p.name]
            tds.append(
                f'<td style="padding:4px 8px 4px 0;">'
                f'<img src="cid:{cid}" width="{width}"></td>'
            )
        rows_html.append("<tr>" + "".join(tds) + "</tr>")
    return '<table border="0" cellspacing="0" cellpadding="0">' + "".join(rows_html) + "</table>"


def create_message():
    # 收集文件
    pngs, top_imgs, term_imgs, excels = collect_files(FOLDER)
    if not pngs:
        raise RuntimeError("指定文件夹中没有找到 PNG 图片，请检查 FOLDER 路径。")

    # 根邮件：multipart/related（为了内嵌图片）
    msg_root = MIMEMultipart('related')
    msg_root['From'] = FROM_ADDR
    msg_root['To'] = ", ".join(TO_ADDRS)
    if CC_ADDRS:
        msg_root['Cc'] = ", ".join(CC_ADDRS)
    msg_root['Subject'] = Header(SUBJECT, 'utf-8')

    # alternative：只有 HTML（如需要可加纯文本）
    msg_alt = MIMEMultipart('alternative')
    msg_root.attach(msg_alt)

    # 先占好 CID 映射
    cid_map = {}
    for i, p in enumerate(pngs):
        cid_map[p.name] = f"img{i + 1}"

    # 用 CID 拼 HTML
    table_top = build_table(top_imgs, cid_map, cols=2, width=480)
    table_term = build_table(term_imgs, cid_map, cols=2, width=480)

    html_body = f"""
    <html>
      <body>
        <p>{GREETING}</p>
        <p>附件及正文为 <b>{DATE_STR} 科创债ETF成分券估值相关信息</b>，供参考。</p>

        <p><b>1. 科创债ETF成分券收益率曲线</b></p>
        {table_top}
        <br/>

        <p><b>2. 各关键期限科创债ETF成分券历史收益率</b></p>
        {table_term}
        <br/>

        <p>{SIGNATURE.replace(chr(10), "<br/>")}</p>
      </body>
    </html>
    """

    msg_alt.attach(MIMEText(html_body, 'html', 'utf-8'))

    # 把图片作为内嵌资源添加（inline）
    for p in pngs:
        with open(p, 'rb') as f:
            img_data = f.read()
        img = MIMEImage(img_data)
        cid = cid_map[p.name]
        img.add_header('Content-ID', f'<{cid}>')
        img.add_header('Content-Disposition', 'inline', filename=p.name)
        msg_root.attach(img)

    # 添加 Excel 附件
    for x in excels:
        ctype, encoding = mimetypes.guess_type(x.name)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)

        with open(x, 'rb') as f:
            part = MIMEBase(maintype, subtype)
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=x.name)
        msg_root.attach(part)

    return msg_root


def send_mail(msg):
    all_recipients = TO_ADDRS + CC_ADDRS

    if USE_SSL:
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT)
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)

    try:
        server.ehlo()
        if not USE_SSL:
            # 如果是 587 端口，通常需要 STARTTLS
            server.starttls()
            server.ehlo()

        server.login(USERNAME, PASSWORD)
        server.sendmail(FROM_ADDR, all_recipients, msg.as_string())
        print("邮件发送成功")
    finally:
        server.quit()


if __name__ == "__main__":
    message = create_message()
    send_mail(message)
