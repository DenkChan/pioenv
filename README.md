# pioenv
# 需先安装依赖：pip install pandas openpyxl
import pandas as pd

try:
    # 读取Excel文件
    df = pd.read_excel('3.1x.xx.xlsx')
    
    # 定义需要提取的列名
    target_columns = ['CAR_TYPE', '填充1', '填充2', '封装方式', '代号', '版本号', '日期版本', '修改次数', '是否合成']
    
    # 验证所有目标列是否存在
    missing_columns = [col for col in target_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel文件中缺少以下必要列: {', '.join(missing_columns)}")
    
    # 提取目标列数据并转换为列表
    result = df[target_columns].values.tolist()
    
    for cartype in result['CAR_TYPE']:
        subprocess.run(['gpjmodify','set',':macro=CAR_TYPE={}'.format(cartype)],':macro=SOFT_VERSION={}'.format(),':macro=DATE_VERSION=2023-10-01',':macro=MODIFY_COUNT=1',':macro=IS_SYNTHESIZED=true'])
    # 打印结果（可根据需要修改输出方式）
    print("提取结果:\n", result)
    
except FileNotFoundError:
    print("错误: 未找到'3.1x.xx.xlsx'文件，请确保文件在当前目录下")
except Exception as e:
    print(f"处理过程中出错: {str(e)}")
