import pandas as pd


def get_college_name(dict, serial_number):
    return dict.get(serial_number, "序号不存在")


file_path = './files/CollegeNum2Name.xlsx'
df = pd.read_excel(file_path, header=None)

data = df.iloc[1:19, [0, 2]]
data.columns = ['序号', '学院名称']

college_dict = dict(zip(data['序号'], data['学院名称']))
