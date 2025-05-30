import requests
import time
import pandas as pd
from datetime import datetime, timedelta

headers = {
    "accept": "application/json, text/javascript, */*; q=0.01",
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

url = "https://webapi.sporttery.cn/gateway/uniform/football/getUniformMatchResultV1.qry"


def validate_date(date_str):
    try:
        datetime.strptime(date_str, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def get_date_range():
    while True:
        start_date = input("请输入开始日期（格式：YYYY-MM-DD）：")
        if validate_date(start_date):
            break
        print("日期格式错误，请重新输入！")

    while True:
        end_date = input("请输入结束日期（格式：YYYY-MM-DD）：")
        if validate_date(end_date):
            break
        print("日期格式错误，请重新输入！")

    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")

    if start > end:
        start, end = end, start
        start_date, end_date = end_date, start_date  # 保持原始输入顺序

    date_list = []
    current_date = start
    while current_date <= end:
        date_list.append(current_date.strftime("%Y-%m-%d"))
        current_date += timedelta(days=1)

    return date_list, start_date, end_date


def get_daily_matches(target_date):
    params = {
        "matchBeginDate": target_date,
        "matchEndDate": target_date,
        "leagueId": "",
        "pageSize": "30",
        "pageNo": "1",
        "isFix": "0",
        "matchPage": "1",
        "pcOrWap": "1"
    }

    daily_matches = []
    current_page = 1
    page_size = int(params['pageSize'])

    while True:
        try:
            params['pageNo'] = str(current_page)
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()

            data = response.json()
            if not data.get('success'):
                print(f"{target_date} 第 {current_page} 页请求失败：{data.get('message')}")
                break

            matches = data.get('value', {}).get('matchResult', [])
            if not matches:
                print(f"{target_date} 第 {current_page} 页没有数据")
                break

            daily_matches.extend(matches)
            print(f"{target_date} 第 {current_page} 页获取 {len(matches)} 条数据")

            if len(matches) < page_size:
                break

            current_page += 1
            time.sleep(0.5)

        except Exception as e:
            print(f"发生异常：{str(e)}")
            break

    return daily_matches


def main():
    date_range, start_date, end_date = get_date_range()
    all_matches = []

    print("\n开始爬取数据...")
    for date in date_range:
        print(f"\n正在获取 {date} 的数据...")
        daily_data = get_daily_matches(date)
        all_matches.extend(daily_data)
        print(f"{date} 共获取 {len(daily_data)} 条数据")
        time.sleep(1)

    print("\n最终统计：")
    print(f"总天数：{len(date_range)} 天")
    print(f"总比赛场次：{len(all_matches)} 场")

    if all_matches:
        # 转换为DataFrame
        df = pd.DataFrame([{
            '场次编码': match.get('matchId', 'N/A'),
            '赛事日期': match.get('matchDate', '无日期'),
            '赛事编号': match.get('matchNumStr', '无编号'),
            '联赛名称': match.get('leagueNameAbbr', '未知联赛'),
            '主队': match.get('allHomeTeam', '未知主队'),
            '客队': match.get('allAwayTeam', '未知客队')
        } for match in all_matches])

        # 生成文件名
        filename = f"比赛数据_{start_date}_至_{end_date}.xlsx"

        try:
            # 保存Excel文件
            df.to_excel(filename, index=False, engine='openpyxl')
            print(f"\n数据已成功保存到：{filename}")
            print("文件包含以下字段：")
            print(df.columns.tolist())
        except Exception as e:
            print(f"\n保存文件时出错：{str(e)}")
            print("正在尝试保存为CSV格式...")
            try:
                csv_filename = f"比赛数据_{start_date}_至_{end_date}.csv"
                df.to_csv(csv_filename, index=False)
                print(f"数据已成功保存到：{csv_filename}")
            except Exception as e2:
                print(f"CSV保存也失败：{str(e2)}")

        # 打印前5条数据预览
        print("\n数据预览：")
        print(df.head())
    else:
        print("没有需要保存的数据")


if __name__ == "__main__":
    main()