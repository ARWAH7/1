import requests
import pandas as pd
from datetime import datetime

headers = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "origin": "https://www.sporttery.cn",
    "priority": "u=1, i",
    "referer": "https://www.sporttery.cn/",
    "sec-ch-ua": "\"Chromium\";v=\"136\", \"Microsoft Edge\";v=\"136\", \"Not.A/Brand\";v=\"99\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
}
url = "https://webapi.sporttery.cn/gateway/uniform/football/getFixedBonusV1.qry"
params = {"clientCode": "3001", "matchId": "1028935"}


def fetch_and_process_data():
    response = requests.get(url, headers=headers, params=params)
    data = response.json()

    # 提取单关标识信息
    single_had = "否"
    for item in data.get('value', {}).get('oddsHistory', {}).get('singleList', []):
        if item.get('poolCode') == 'HAD' and str(item.get('single', '')) == '1':
            single_had = "是"
            break

    # 提取赛事结果
    result_mapping = {}
    for result in data.get('value', {}).get('matchResultList', []):
        code = result.get('code', '')
        desc = result.get('combinationDesc', '')
        if code == 'HAD':
            result_mapping['胜平负结果'] = desc
        elif code == 'HHAD':
            result_mapping['让球结果'] = desc
        elif code == 'CRS':
            result_mapping['比分结果'] = desc
        elif code == 'TTG':
            result_mapping['总进球结果'] = desc
        elif code == 'HAFU':
            result_mapping['半全场结果'] = desc

    # 提取所有赔率数据
    had_list = data.get('value', {}).get('oddsHistory', {}).get('hadList', [])
    hhad_list = data.get('value', {}).get('oddsHistory', {}).get('hhadList', [])
    ttg_list = data.get('value', {}).get('oddsHistory', {}).get('ttgList', [])
    hafu_list = data.get('value', {}).get('oddsHistory', {}).get('hafuList', [])
    crs_list = data.get('value', {}).get('oddsHistory', {}).get('crsList', [])

    # 合并数据
    max_length = max(len(had_list), len(hhad_list), len(ttg_list), len(hafu_list), len(crs_list))
    combined = []

    for i in range(max_length):
        had = had_list[i] if i < len(had_list) else {}
        hhad = hhad_list[i] if i < len(hhad_list) else {}
        ttg = ttg_list[i] if i < len(ttg_list) else {}
        hafu = hafu_list[i] if i < len(hafu_list) else {}
        crs = crs_list[i] if i < len(crs_list) else {}

        record = {
            # ========== 基础信息 ==========
            'matchId': params['matchId'],
            '单关标识': single_had,

            # ========== 胜平负数据 ==========
            '胜平负发布日期': had.get('updateDate', ''),
            '胜平负发布时间': had.get('updateTime', ''),
            '胜': had.get('h', ''),
            '平': had.get('d', ''),
            '负': had.get('a', ''),

            # ========== 让球数据 ==========
            '让球发布日期': hhad.get('updateDate', ''),
            '让球发布时间': hhad.get('updateTime', ''),
            '让球数': hhad.get('goalLine', ''),
            '让球胜': hhad.get('h', ''),
            '让球平': hhad.get('d', ''),
            '让球负': hhad.get('a', ''),

            # ========== 总进球数据 ==========
            '进球数发布日期': ttg.get('updateDate', ''),
            '进球数发布时间': ttg.get('updateTime', ''),
            '0球': ttg.get('s0', ''),
            '1球': ttg.get('s1', ''),
            '2球': ttg.get('s2', ''),
            '3球': ttg.get('s3', ''),
            '4球': ttg.get('s4', ''),
            '5球': ttg.get('s5', ''),
            '6球': ttg.get('s6', ''),
            '7球': ttg.get('s7', ''),

            # ========== 半全场数据 ==========
            '半全场发布日期': hafu.get('updateDate', ''),
            '半全场发布时间': hafu.get('updateTime', ''),
            '胜胜': hafu.get('hh', ''),
            '胜平': hafu.get('hd', ''),
            '胜负': hafu.get('ha', ''),
            '平胜': hafu.get('dh', ''),
            '平平': hafu.get('dd', ''),
            '平负': hafu.get('da', ''),
            '负胜': hafu.get('ah', ''),
            '负平': hafu.get('ad', ''),
            '负负': hafu.get('aa', ''),

            # ========== 比分数据 ==========
            '比分发布日期': crs.get('updateDate', ''),
            '比分发布时间': crs.get('updateTime', ''),
            '1:0': crs.get('s01s00', ''),
            '2:0': crs.get('s02s00', ''),
            '2:1': crs.get('s02s01', ''),
            '3:0': crs.get('s03s00', ''),
            '3:1': crs.get('s03s01', ''),
            '3:2': crs.get('s03s02', ''),
            '4:0': crs.get('s04s00', ''),
            '4:1': crs.get('s04s01', ''),
            '4:2': crs.get('s04s02', ''),
            '5:0': crs.get('s05s00', ''),
            '5:1': crs.get('s05s01', ''),
            '5:2': crs.get('s05s02', ''),
            '胜其它': crs.get('s-1sh', ''),
            '0:0': crs.get('s00s00', ''),
            '1:1': crs.get('s01s01', ''),
            '2:2': crs.get('s02s02', ''),
            '3:3': crs.get('s03s03', ''),
            '平其它': crs.get('s-1sd', ''),
            '0:1': crs.get('s00s01', ''),
            '0:2': crs.get('s00s02', ''),
            '1:2': crs.get('s01s02', ''),
            '0:3': crs.get('s00s03', ''),
            '1:3': crs.get('s01s03', ''),
            '2:3': crs.get('s02s03', ''),
            '0:4': crs.get('s00s04', ''),
            '1:4': crs.get('s01s04', ''),
            '2:4': crs.get('s02s04', ''),
            '0:5': crs.get('s00s05', ''),
            '1:5': crs.get('s01s05', ''),
            '2:5': crs.get('s02s05', ''),
            '负其它': crs.get('s-1sa', ''),

            # ========== 赛事结果 ==========
            '胜平负结果': result_mapping.get('胜平负结果', ''),
            '让球结果': result_mapping.get('让球结果', ''),
            '比分结果': result_mapping.get('比分结果', ''),
            '总进球结果': result_mapping.get('总进球结果', ''),
            '半全场结果': result_mapping.get('半全场结果', '')
        }
        combined.append(record)

    return combined


# 完整列定义
columns = [
    'matchId', '单关标识',
    '胜平负发布日期', '胜平负发布时间', '胜', '平', '负',
    '让球发布日期', '让球发布时间', '让球数', '让球胜', '让球平', '让球负',
    '进球数发布日期', '进球数发布时间',
    '0球', '1球', '2球', '3球', '4球', '5球', '6球', '7球',
    '半全场发布日期', '半全场发布时间',
    '胜胜', '胜平', '胜负', '平胜', '平平', '平负', '负胜', '负平', '负负',
    '比分发布日期', '比分发布时间',
    '1:0', '2:0', '2:1', '3:0', '3:1', '3:2',
    '4:0', '4:1', '4:2', '5:0', '5:1', '5:2',
    '胜其它', '0:0', '1:1', '2:2', '3:3', '平其它',
    '0:1', '0:2', '1:2', '0:3', '1:3', '2:3',
    '0:4', '1:4', '2:4', '0:5', '1:5', '2:5', '负其它',
    '胜平负结果', '让球结果', '比分结果', '总进球结果', '半全场结果'
]

df = pd.DataFrame(fetch_and_process_data(), columns=columns)

# 保存到Excel
with pd.ExcelWriter('完整足球数据.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='综合数据')

    # 设置格式
    workbook = writer.book
    worksheet = writer.sheets['综合数据']

    # 通用数字格式
    number_format = workbook.add_format({'num_format': '0.00'})
    number_columns = [
        ('E:G', 8),  # 胜平负
        ('J:L', 8),  # 让球
        ('O:V', 8),  # 总进球
        ('Y:AG', 8),  # 半全场
        ('AJ:BR', 8)  # 比分
    ]
    for cols, width in number_columns:
        worksheet.set_column(cols, width, number_format)

    # 时间格式
    time_format = workbook.add_format({'num_format': 'hh:mm:ss'})
    time_columns = [
        ('D:D', 12),  # 胜平负时间
        ('I:I', 12),  # 让球时间
        ('P:P', 12),  # 总进球时间
        ('X:X', 12),  # 半全场时间
        ('AI:AI', 12)  # 比分时间
    ]
    for cols, width in time_columns:
        worksheet.set_column(cols, width, time_format)

    # 日期格式
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    date_columns = [
        ('C:C', 14),  # 胜平负日期
        ('H:H', 14),  # 让球日期
        ('O:O', 14),  # 总进球日期
        ('W:W', 14),  # 半全场日期
        ('AH:AH', 14)  # 比分日期
    ]
    for cols, width in date_columns:
        worksheet.set_column(cols, width, date_format)

    # 特殊格式
    worksheet.set_column('J:J', 8, workbook.add_format({'num_format': '0.00;-0.00'}))  # 让球数
    worksheet.set_column('A:A', 12, workbook.add_format({'num_format': '@'}))  # matchId文本格式
    worksheet.set_column('B:B', 10)  # 单关标识列宽

print("数据文件已生成：完整足球数据.xlsx")