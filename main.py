from xml.etree import ElementTree as ET
import xlsxwriter

"""
这是为盐和避难所1010版本编写的怪物信息提取脚本
本脚本使用两个文件作为输入
其文件名应分别为"monster-enhanced.xml"
和"monster-classic.xml"
本脚本的输出是一个excel表，该表会列出各种怪物的血量、抗性等信息
"""

classic_file = "data/monsters-classic.xml"
enhanced_file = "data/monsters-enhanced.xml"

'''
下述两个字典的组织形式：
[{name:, def1:, def2:, def3:, hp:,}, ]
'''
classic_info = []
enhanced_info = []
infos = ['name',
         'title',
         'hp',
         'defense',
         'poise',
         'stamina',
         'fireDef',
         'litDef',
         'bladedDef',
         'poisonDef',
         'holyDef',
         'darkDef']


def parse_monster(file_name, stored_info: list):
    tree = ET.parse(file_name)
    MonstersEditor = tree.getroot()
    monsters = MonstersEditor[0]
    for Monsters in monsters:
        monster = {}
        for info in infos:
            node = Monsters.find(info)
            if info == 'title':
                monster[info] = node[7].text
            else:
                monster[info] = node.text
        stored_info.append(monster)


write_seq = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']


def write_to_excel(file_name, stored_info: list):
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    write_index = 1
    for i in range(0, 12):
        worksheet.write(write_seq[i] + str(write_index), infos[i])
    write_index += 1
    for info in stored_info:
        for i in range(0, 12):
            worksheet.write(write_seq[i] + str(write_index), info[infos[i]])
        write_index += 1
    workbook.close()


parse_monster(classic_file, classic_info)
parse_monster(enhanced_file, enhanced_info)
merged_info = []
to_merge = ['hp',
         'defense',
         'poise',
         'stamina',
         'fireDef',
         'litDef',
         'bladedDef',
         'poisonDef',
         'holyDef',
         'darkDef']
for i in range(len(classic_info)):
    classic = classic_info[i]
    enhanced = enhanced_info[i]
    if classic['name'] != enhanced['name']:
        print('bad!')
    merged = {}
    merged['name'] = classic['name']
    merged['title'] = classic['title']
    for k in to_merge:
        value1 = classic[k]
        value2 = enhanced[k]
        merged[k] = str(value1) + '/' + str(value2)
    merged_info.append(merged)


# 写入经典怪物数据
write_to_excel(classic_file.replace('xml', 'xlsx'), classic_info)
# 写入加强版怪物数据
write_to_excel(enhanced_file.replace('xml', 'xlsx'), enhanced_info)
# 写入合并后的怪物数据
write_to_excel('data/merged.xlsx', merged_info)


