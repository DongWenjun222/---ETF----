import sys
import subprocess

subprocess.check_call([
    sys.executable, "-m", "pip", "install",
    "-r", "requirements.txt",
    "--upgrade-strategy", "only-if-needed",
])

import time
from pathlib import Path
import xlwings as xw

xlCalcStateDone = 0

def wait_until_done(app, timeout=600):
    """等待 Excel 完成所有计算/异步查询（含 PowerQuery），最多 timeout 秒"""
    start = time.time()
    while True:
        try:
            if app.api.CalculationState == xlCalcStateDone and getattr(app.api, "Ready", True):
                break
        except Exception:
            if app.api.CalculationState == xlCalcStateDone:
                break
        if time.time() - start > timeout:
            print("[警告] 等待计算超时，继续后续步骤")
            break
        
    # 再尝试等待异步查询（Office 2016+）
    try:
        for _ in range(3):
            res = app.api.CalculateUntilAsyncQueriesDone()
            if res is True:
                break
    except Exception:
        pass

def disable_background_refresh(wb):
    """尽量关闭后台刷新，避免 RefreshAll 立即返回导致保存半成品"""
    try:
        for conn in wb.api.Connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                if hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass
    except Exception:
        pass
    # Worksheet 级的 QueryTable & ListObject
    try:
        for ws in wb.sheets:
            try:
                for qt in ws.api.QueryTables:
                    qt.BackgroundQuery = False
            except Exception:
                pass
            try:
                for lo in ws.api.ListObjects:
                    try:
                        qt = lo.QueryTable
                        if qt:
                            qt.BackgroundQuery = False
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

def refresh_folder(folder, pattern="*.xls*"):
    folder = Path(folder)
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        # 关键改动：不再设置 app.api.Calculation（避免你的报错）
        for f in folder.glob(pattern):
            
            if f.name.startswith("~$") or not f.is_file():
                continue
            print(f"[打开] {f.name}")
            wb = app.books.open(f, update_links=False, read_only=False)
            try:
                # 保险：尽量关闭后台刷新
                disable_background_refresh(wb)

                # 强制全量重算，能触发 Wind XLL 公式（WSD/WSS/WST 等）
                print(f"[重算] {f.name} -> Application.CalculateFullRebuild()")
                app.api.CalculateFullRebuild()
                wait_until_done(app)

                # 如有 PowerQuery/外部连接，再统一刷新
                try:
                    print(f"[刷新连接] {f.name} -> Workbook.RefreshAll()")
                    wb.api.RefreshAll()
                    wait_until_done(app)
                except Exception as e:
                    print(f"[提示] RefreshAll 出现异常：{e}")

                time.sleep(60)
                wb.save()
                print(f"[完成] {f.name} 已保存")
            finally:
                wb.close()
    finally:
        app.quit()
        print("[退出] Excel 实例已关闭")

if __name__ == "__main__":
    # 直接把你的路径放进来即可
    refresh_folder(r".\\", "*更新.xlsx")
