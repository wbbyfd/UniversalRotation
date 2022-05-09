import os, random, time
import xlwings as xw
import pandas as pd
import requests
import pysnowball

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

@xw.func
def get_fund_net_asset_value_history(fund_code: str, pz: int = 30) -> pd.DataFrame:
    '''
    根据基金代码和要获取的页码抓取基金净值信息

    Parameters
    ----------
    fund_code : 6位基金代码
    page : 页码 1 为最新页数据

    Return
    ------
    DataFrame : 包含基金历史k线数据
    '''
    # 请求头
    EastmoneyFundHeaders = {
        'Accept': '*/*',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        # 'Connection': 'keep-alive',
        'Host': 'api.fund.eastmoney.com',
        'Referer': 'http://fundf10.eastmoney.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36',
    }
    # 请求参数
    Eastmoneyparams = {
        'fundCode': f'{fund_code}',
        'pageIndex': '1',
        'pageSize': f'{pz}',
        'startDate': '',
        'endDate': '',
        '_': round(time.time()*1000),
    }
    EastmoneyCookie = {
        'qgqp_b_id': "e1dc92ae3b0d76cc57233f6d6a1c5c61",
    }
    url = 'http://api.fund.eastmoney.com/f10/lsjz'

    #设置重连次数
    requests.adapters.DEFAULT_RETRIES = 10
    session1 = requests.session()
    # 设置连接活跃状态为False
    session1.keep_alive = False
    response1 = session1.get(url, headers=EastmoneyFundHeaders, params=Eastmoneyparams, cookies=EastmoneyCookie, verify=False, stream=False, timeout=10)
    json_response = response1.json()
    response1.close()
    del(response1)

    rows = []
    columns = ['日期', '单位净值', '累计净值', '涨跌幅']
    if json_response is None:
        return pd.DataFrame(rows, columns=columns)
    datas = json_response['Data']['LSJZList']
    if len(datas) == 0:
        return pd.DataFrame(rows, columns=columns)
    for stock in datas:
        rows.append({
            '日期': stock['FSRQ'],
            '单位净值': stock['DWJZ'],
            '累计净值': stock['LJJZ'],
            '涨跌幅': stock['JZZZL']
        })

    df = pd.DataFrame(rows)
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
    df['单位净值'] = pd.to_numeric(df['单位净值'], errors='coerce')
    df['累计净值'] = pd.to_numeric(df['累计净值'], errors='coerce')
    df['涨跌幅'] = pd.to_numeric(df['涨跌幅'], errors='coerce')
    return df

@xw.func
# 获取基金溢价率
def get_fund_premium_rate(fund_code: str):
    detail = pd.DataFrame(pysnowball.quote_detail(fund_code))
    row1=detail.loc["quote"]
    col1=row1[0]["premium_rate"]
    return col1

@xw.func
# 根据排名计算排名分(精确到小数点后2位)
def get_rank_value(length: int, rank: int):
    return round((length-rank)/length*100, 2)

@xw.func
# 根据20天净值增长率和溢价率来量化轮动基金
def rotate_fund_by_premium_rate_and_20net_asset_value(source_sheets: str, source_range: str, dest_range: str, delay: int):
    xw.Book("UniversalRotation.xlsm").set_mock_caller()
    wb = xw.Book.caller()  # wb = xw.Book(r'UniversalRotation.xlsm')
    pd.options.display.max_columns = None
    pd.options.display.max_rows = None
    pysnowball.set_token('xq_a_token=e8119f7d7a050cdbfa822fa0da4de5bec1ee0dc7;')

    sheet_fund = wb.sheets[source_sheets]
    data_fund = pd.DataFrame(sheet_fund.range(source_range).value,
                                       columns=['基金代码','基金名称','投资种类','20天净值增长率','溢价率','总排名分',
                                                '20天净值增长率_排名分','溢价率_排名分'])
    log_file = open('log-' + source_sheets + str(time.strftime("-%Y-%m-%d-%H-%M-%S", time.localtime())) + '.txt', 'a+')

    for i,fund_code in enumerate(data_fund['基金代码']):
        # 更新溢价率
        fundPremiumRateValue = get_fund_premium_rate(fund_code)
        data_fund.loc[i, '溢价率'] = fundPremiumRateValue
        # 延时
        if (i % 10 == 0) :
            time.sleep(random.randrange(delay))

        netAssetValue = get_fund_net_asset_value_history(fund_code[2:8])
        # 最新净值
        netAssetDate1 = netAssetDate20 = str(netAssetValue.loc[0][0])[0:10]
        netAssetValue1 = netAssetValue20 = netAssetValue.loc[0][2]
        # 20个交易日前的净值
        if (len(netAssetValue) > 20) :
            netAssetDate20 = str(netAssetValue.loc[20][0])[0:10]
            netAssetValue20 = netAssetValue.loc[20][2]
        # 更新20天净值增长率
        netAssetValue20Rate = round((netAssetValue1 - netAssetValue20) / netAssetValue20 * 100, 2)
        data_fund.loc[i, '20天净值增长率'] = netAssetValue20Rate

        log_str = 'No.' + format(str(i), "<6") + fund_code + ':'+ format(data_fund.loc[i, '基金名称'], "<15") \
                  + netAssetDate1 + ':净值:' + format(str(netAssetValue1), "<10") \
                  + netAssetDate20 + ':净值:' + format(str(netAssetValue20), "<10") \
                  + '二十个交易日净值增长率:' + format(str(netAssetValue20Rate), "<10") + '溢价率:'+ format(str(fundPremiumRateValue), "<10")
        print(log_str)
        print(log_str, file=log_file)

    # 更新'溢价率_排名分'
    data_fund = data_fund.sort_values(by='溢价率', ascending=True)
    data_fund.reset_index(drop=True, inplace=True)
    for i,fundPremiumRateValue in enumerate(data_fund['溢价率']):
        # print('溢价率：' + str(fundPremiumRateValue))
        data_fund.loc[i, '溢价率_排名分'] = get_rank_value(len(data_fund), i)

    # 更新'20天净值增长率_排名分'
    data_fund = data_fund.sort_values(by='20天净值增长率', ascending=False)
    data_fund.reset_index(drop=True, inplace=True)
    for i,netAssetValue in enumerate(data_fund['20天净值增长率']):
        # print('20天净值增长率：' + str(netAssetValue))
        data_fund.loc[i, '20天净值增长率_排名分'] = get_rank_value(len(data_fund), i)

    # 更新'总排名分'
    for i,sumValue in enumerate(data_fund['总排名分']):
        data_fund.loc[i, '总排名分'] = data_fund.loc[i, '溢价率_排名分'] +\
                                           data_fund.loc[i, '20天净值增长率_排名分']
        # print('总排名分'+str(data_fund.loc[i, '总排名分']))
    data_fund = data_fund.sort_values(by='总排名分', ascending=False)
    data_fund.reset_index(drop=True, inplace=True)
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()

    # 更新原Excel
    sheet_fund.range(dest_range).value = data_fund

    wb.save()


@xw.func
# 轮动20天净值增长和溢价率选LOF、ETF和封基
def rotate_LOF_ETF():
    print("------------------------20天净值增长率和溢价率轮动LOF、ETF和封基----------------------------------------------------")
    # 数据区域为'F3:M665'
    rotate_fund_by_premium_rate_and_20net_asset_value('20天净值增长率和溢价率轮动LOF、ETF和封基','F3:M665','E2', 10)

@xw.func
# 轮动20天净增和溢价率选债券和境外基金
def rotate_abroad_fund():
    print("-----------------------20天净值增长率和溢价率轮动债券和境外基金---------------------------------------------------------")
    # 数据区域为'F3:M131'
    rotate_fund_by_premium_rate_and_20net_asset_value('20天净值增长率和溢价率轮动债券和境外基金', 'F3:M131', 'E2', 10)

def main():

    #删除旧log
    for eachfile in os.listdir('./'):
        filename = os.path.join('./', eachfile)
        if os.path.isfile(filename) and filename.startswith("./log") :
            os.remove(filename)

    rotate_LOF_ETF()
    rotate_abroad_fund()

if __name__ == "__main__":
    main()