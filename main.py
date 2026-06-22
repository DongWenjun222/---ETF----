from step1刷新数据 import refresh_folder
from step2合并数据 import merge_data
from step4发送邮件 import create_message,send_mail
import sys
import subprocess                                                                                                                                                                                                                                                                                                                                                                                                                            m'm'm'm'm'm'm'm'm'm'm'm'm'm'm'm'm'm'm

### 运行step1刷新数据_更新数据 
#refresh_folder(r".\\", "*更新.xls*")

### 运行step2合并数据_合并数据
merge_data()

### 运行step3画图.py_生成图表
subprocess.check_call([sys.executable, "step3画图.py"])

### 运行step4发送邮件_发送邮件
message = create_message()
send_mail(message)
