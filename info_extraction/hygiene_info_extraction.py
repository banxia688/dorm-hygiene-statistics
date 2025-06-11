from datetime import datetime

month_abbr = datetime.now().strftime("%b")
year = str(datetime.now().year)
file_path = './files/dorm_hygiene_info/' + year + ' ' + month_abbr + ' Info.txt'
hygiene_data = []
with open(file_path, 'r', encoding='utf-8') as file:
    for line in file:
        # 去掉行尾的换行符，然后按"_"分割
        parts = line.strip().split("_")
        if len(parts) > 3:
            # 将分割后的数据添加到列表中
            hygiene_data.append(parts)

            # 检查最后一列是否包含 "&" 并且包含当前行的学院
            if "&" in parts[3] and parts[0] in parts[3]:
                # 提取 "&" 后面的院系名称
                additional_institute = parts[3].split("&")[1]

                # 生成新行数据
                new_row = [additional_institute] + parts[1:]  # 用 additional_institute 替换第一列
                hygiene_data.append(new_row)

# 进一步处理第三列和最后一列
for entry in hygiene_data:
    # 分割第三列
    entry[2] = entry[2].split("-")
    # 标注混合宿舍
    entry[3] = "混合宿舍" if entry[3] != '无' else ''
