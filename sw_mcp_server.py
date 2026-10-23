import win32com.client
from fastmcp import FastMCP
import sw_sim_api

# 连接SolidWorks，提前手动打开SW！
sw = win32com.client.Dispatch("SldWorks.Application")
sw.Visible = True

server = FastMCP("SW自动仿真服务")

# 工具1：运行SW建模脚本
@server.tool()
def sw_run_script(script: str):
    exec(script)
    return "✅建模脚本执行成功"

# 工具2：读取仿真结果
@server.tool()
def sw_get_simulation_result():
    res = sw_sim_api.get_sim_result(sw)
    judge = sw_sim_api.check_result(res, "")
    return {"仿真数据":res,"判定":judge}

# 工具3：导出模态振型动画
@server.tool()
def sw_export_mode_animation(case_name:str, mode_order:int, save_path:str):
    ret = sw_sim_api.export_mode_animation(sw, case_name, mode_order, save_path)
    return ret

# ========== 新增工具4：静力学应力变形动画导出 ==========
@server.tool()
def sw_export_static_stress_animation(case_name:str, save_path:str):
    ret = sw_sim_api.export_static_stress_animation(sw, case_name, save_path)
    return ret

# ========== 新增工具5：生成Word仿真报告 ==========
@server.tool()
def sw_create_simulation_word_report(report_title:str, sim_data:dict, screenshot_path:str, save_docx_path:str):
    ret = sw_sim_api.create_simulation_word_report(sw, report_title, sim_data, screenshot_path, save_docx_path)
    return ret


if __name__ == "__main__":
    server.run()