from util import get_specific_letter_by_increment


header = [
    'name',
    '名称',
    '类别',
    '强化素材',
    '强化系数',
    'flags',
    'special',
    '重量',
    '基础面,板',
    '力量补正',
    '敏捷补正',
    '智慧补正',
    '魔力补正',
    '武器牌照',
    '第二牌照',
    '魔力/智力补正（仅武器）',
    '变质材料',
    '需要数量',
    '变质介质',
    '需要盐量'
    '价格',
    '耐久',
]

values = [0, 1, 2, 3, 4, 5, 6, 9]

upgrade = []


def write_header(worksheet):
    i = 0
    for h in header:
        worksheet.write(get_specific_letter_by_increment(i) + '1', h)
        i += 1


def write_data(data: dict, worksheet, write_index):
    i = 0
    for k in data.keys():
        if k == 'values':
            values_ = data[k]
            for writable_value_index in values:
                worksheet.write(get_specific_letter_by_increment(i) + write_index, values_[writable_value_index])
        elif k == 'upgradePath':
            writeable_upgrade_path = data[k][0]
            for k_ in writeable_upgrade_path.keys():
                worksheet.write(get_specific_letter_by_increment(i) + write_index, writeable_upgrade_path[k_])
        else:
            worksheet.write(get_specific_letter_by_increment(i) + write_index, data[k])


def write_to_excel(weapon_data: list, worksheet):
    write_index = 2
    for wd in weapon_data:
        write_data(wd, worksheet, write_index)
        write_index += 1