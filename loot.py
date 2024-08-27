from xml.etree import ElementTree as ET
import xlsxwriter

from util import get_specific_letter_by_increment
import weapon

"""
这是为盐和避难所1010版本编写的武器、凋落物等信息提取脚本
本脚本使用两个文件作为输入
其文件名应分别为"loot-enhanced.xml"
和"loot-classic.xml"
本脚本的输出是一个excel表，该表会列出各种武器的信息
"""

classic_file = "loot-classic.xml"
enhanced_file = "loot-enhanced.xml"

'''
提取的信息：
category：
type：
upgrade：
upgradeFac：
name：
title：
flags：
special：
weight：
values：[float, ]
upgradePath：[reqLoot:, outLoot:, cost:, reqCount:,]
value：
durability:
texIdx:
texIdx2:
'''
classic_info = {}
enhanced_info = {}
categories = {
    0: '武器',
    1: '盾牌',
    2: '衣服',
    3: '戒指',
    4: '消耗品',
    5: '法术',
    6: '钥匙',
    7: '强化素材'
}
for cat in categories.keys():
    classic_info[cat] = list()
    enhanced_info[cat] = list()

types = {
    0: {
        0: 'dagger/匕首',
        1: 'shortsword/短剑',
        2: 'mace/锤',
        3: 'axe/斧',
        4: 'spear/枪/戟/叉',
        5: 'scythe/镰',
        6: 'claymore/大剑',
        7: 'hammer/大锤',
        8: 'bow/弓',
        9: 'crossbow/弩',
        10: 'staff/杖',
        11: 'whip/鞭',
        12: 'axe/巨斧',
        13: 'halberd/长柄斧',
        14: 'gun/枪',
        15: 'torch/火把',
        16: 'wand/枝',
        17: 'sword/巨剪',
        18: 'sword_gun/枪刃',
        19: 'sword_whip/剑鞭',
        20: 'sword/单手剑',
    },
    11: {
        1: '小盾',
        5: '圆盾',
        3: '大盾/塔盾',
        4: '砍盾'
    },
    10: {
        0: '小盾',
        1: '中盾',
        3: '图纹盾',
        4: '砍盾'
    },
    2: {
        0: 'helm/头盔',
        1: 'armor/胸甲/铠甲',
        2: 'gloves/手套/臂铠',
        3: 'boots/裤/腿甲'
    },
    3: {
        0: 'ring/指环',
        1: 'charm/护符',
        2: 'rune/烙印',
    },
    4: {
        0: '圣像/盐/教派消耗品/附魔/投掷物',
        1: 'arrow/箭矢',
        2: 'bolt/弩箭',
        3: 'flintshot/弹药'
    },
    5: {
        0: 'lit/fire/mag',
        1: 'flame/lit/rock/mag',
        2: 'dark/暗术',
        3: '暗黑群魔',
        5: 'holy/圣术'
    },
    6: {
        0: 'key/关键道具',
    },
    7: {
        0: 'red/烧焦',
        1: 'blue/冰冻',
        2: 'white/士兵',
        3: 'black/黑色',
        4: 'drowning/浸水',
        5: 'amber/琥珀',
        6: 'beast/野兽',
        7: 'dice/随机掉落'
    },
}

floats = {
    0: {
        0: '基础面板',
        1: '力量补正',
        2: '敏捷补正',
        3: '智力补正',
        4: '魔力补正',
        5: '武器牌照',
        6: '第二牌照',
        7: '',
        8: '',
        9: '魔力/智力补正（仅武器）',
        10: '',
        11: '',
        12: '',
        13: '',
        14: '',
        15: '',
        16: '',
        17: '',
        18: '',
        19: '',
        20: '',
        21: '',
        22: '',
        23: '',
        24: '',
        25: '',
    },
    1: {
        0: '硬直减免',
        1: '',
        2: '打击防御',
        3: '火焰防御',
        4: '闪电防御',
        5: '斩击防御',
        6: '毒药防御',
        7: '神圣防御',
        8: '奥术防御',
        9: '',
        10: '',
        11: '',
        12: '',
    },
    2: {
        0: '打击防御',
        1: '火焰防御',
        2: '闪电防御',
        3: '斩击防御',
        4: '毒药防御',
        5: '神圣防御',
        6: '奥术防御',
        7: '平衡',
        8: '',
        9: '',
        10: '',
        11: '',
    },
    3: {
        0: '',
        1: '',
    },
    4: {
        0: ''
    },
    5: {
        0: '需求智慧等级（猜测）',
        1: '',
        2: '蓝耗',
    }
}

infos = ['name',
         'title',
         'category',
         'type',
         'upgrade',
         'upgradeFac',
         'flags',
         'special',
         'weight',
         'values',
         'upgradePath',
         'value',
         'durability']

translate = {
    'name': '名称',
    'title': '中文名',
    'category': '种类',
    'type': '类型',
    'upgrade': '升级素材',
    'upgradeFac': '强化系数',
    'flags': '标记',
    'special': '特殊',
    'weight': '重量',
    'value': '价值',
    'durability': '耐久',
}

upgradePathInfo = [
    'reqLoot',
    'outLoot',
    'cost',
    'reqCount'
]


def extract_all_tags_of_upgrade_path(file_name):
    uniqueUpgradePathInfo = set()
    tree = ET.parse(file_name)
    LootEditor = tree.getroot()
    category = LootEditor.find('category')
    for LootCategory in category.findall('LootCategory'):
        for loot in LootCategory.findall('loot'):
            for LootItem in loot.findall('LootItem'):
                lootItem = {}
                for info in infos:
                    node = LootItem.find(info)
                    if info == 'upgradePath':
                        lootItem[info] = {}
                        for UpgradePath in node.findall('UpgradePath'):
                            for upInfo in UpgradePath:
                                lootItem[info][upInfo.tag] = upInfo.text
    return uniqueUpgradePathInfo


def parse_loot(file_name, stored_info: dict):
    tree = ET.parse(file_name)
    LootEditor = tree.getroot()
    category = LootEditor.find('category')
    for LootCategory in category.findall('LootCategory'):
        for loot in LootCategory.findall('loot'):
            for LootItem in loot.findall('LootItem'):
                lootItem = {}
                for info in infos:
                    node = LootItem.find(info)
                    if info == 'upgradePath':
                        lootItem[info] = []
                        for UpgradePath in node.findall('UpgradePath'):
                            up = {}
                            for upInfo in UpgradePath:
                                up[upInfo.tag] = upInfo.text
                            lootItem[info].append(up)
                    elif info == 'title':
                        if node[7].text is None:
                            lootItem[info] = node[0].text
                        else:
                            lootItem[info] = node[7].text
                    elif info == 'values':
                        lootItem[info] = []
                        for value in node:
                            lootItem[info].append(value.text)
                    else:
                        lootItem[info] = node.text
                stored_info[int(lootItem['category'])].append(lootItem)
    print(stored_info)

def write_header(worksheet, info: dict):
    i = 0
    for k in info.keys():
        if k == 'values':
            cur_category = int(info['category'])
            for j in range(0, len(info[k])):
                if cur_category in floats.keys():
                    worksheet.write(get_specific_letter_by_increment(i) + '1', floats[cur_category][j])
                    i += 1
            pass
        elif k == 'upgradePath':
            pass
        elif k == 'category':
            pass
        elif k == 'upgrade':
            if info['category'] == str(0) or info['category'] == str(1):
                worksheet.write(get_specific_letter_by_increment(i) + '1', translate[k])
                i += 1
        elif k == 'upgradeFac':
            if info['category'] == str(0) or info['category'] == str(1):
                worksheet.write(get_specific_letter_by_increment(i) + '1', translate[k])
                i += 1
        else:
            worksheet.write(get_specific_letter_by_increment(i) + '1', translate[k])
            i += 1


def write_data(worksheet, info: dict, write_index, classic: bool):
    i = 0
    for k in info.keys():
        if k == 'values':
            cur_category = int(info['category'])
            for j in range(len(info[k])):
                if cur_category in floats.keys():
                    worksheet.write(get_specific_letter_by_increment(i) + write_index, info[k][j])
                    i += 1
            pass
        elif k == 'upgradePath':
            pass
        elif k == 'category':
            pass
        elif k == 'type':
            k_ = int(info['category'])
            if k_ == 1:
                if classic:
                    k_ = 10
                else:
                    k_ = 11
            if str(k_) in types.keys():
                worksheet.write(get_specific_letter_by_increment(i) + write_index, types[k_][int(info[k])])
                i += 1
        elif k == 'upgrade':
            if info['category'] == str(0) or info['category'] == str(1):
                worksheet.write(get_specific_letter_by_increment(i) + write_index, types[7][int(info[k])])
                i += 1
        elif k == 'upgradeFac':
            if info['category'] == str(0) or info['category'] == str(1):
                worksheet.write(get_specific_letter_by_increment(i) + write_index, info[k])
                i += 1
        else:
            worksheet.write(get_specific_letter_by_increment(i) + write_index, info[k])
            i += 1


def write_to_excel(file_name, stored_info: dict, classic: bool):
    workbook = xlsxwriter.Workbook(file_name)
    for cur_cat in categories:
        data_list = stored_info[cur_cat]
        worksheet = workbook.add_worksheet(categories[cur_cat])
        write_header(worksheet, data_list[0])
        write_index = 2
        for data in data_list:
            write_data(worksheet, data, str(write_index), classic)
            write_index += 1
    workbook.close()


parse_loot(classic_file, classic_info)
parse_loot(enhanced_file, enhanced_info)

# parse_loot(classic_file, classic_info)
# parse_loot(enhanced_file, enhanced_info)
# merged_info = []
# to_merge = ['hp',
#          'defense',
#          'poise',
#          'stamina',
#          'fireDef',
#          'litDef',
#          'bladedDef',
#          'poisonDef',
#          'holyDef',
#          'darkDef']
# for i in range(len(classic_info)):
#     classic = classic_info[i]
#     enhanced = enhanced_info[i]
#     if classic['name'] != enhanced['name']:
#         print('bad!')
#     merged = {}
#     merged['name'] = classic['name']
#     merged['title'] = classic['title']
#     for k in to_merge:
#         value1 = classic[k]
#         value2 = enhanced[k]
#         merged[k] = str(value1) + '/' + str(value2)
#     merged_info.append(merged)
#
#
# # 写入经典怪物数据
# write_to_excel(classic_file.replace('xml', 'xlsx'), classic_info, True)
# # 写入加强版怪物数据
# write_to_excel(enhanced_file.replace('xml', 'xlsx'), enhanced_info, False)
# # 写入合并后的怪物数据
# write_to_excel('merged.xlsx', merged_info)


