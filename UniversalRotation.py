import os, random, time
import xlwings
import pandas
import requests
import pysnowball
import browser_cookie3

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

@xlwings.func
def get_fund_net_asset_value_history(fund_code: str, pz: int = 30) -> pandas.DataFrame:
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
    response1 = session1.get(url, headers=EastmoneyFundHeaders, params=Eastmoneyparams, cookies=EastmoneyCookie,
                             verify=False, stream=False, timeout=10)
    json_response = response1.json()
    response1.close()
    del(response1)

    rows = []
    columns = ['日期', '单位净值', '累计净值', '涨跌幅']
    if json_response is None:
        return pandas.DataFrame(rows, columns=columns)
    datas = json_response['Data']['LSJZList']
    if len(datas) == 0:
        return pandas.DataFrame(rows, columns=columns)
    for stock in datas:
        rows.append({
            '日期': stock['FSRQ'],
            '单位净值': stock['DWJZ'],
            '累计净值': stock['LJJZ'],
            '涨跌幅': stock['JZZZL']
        })

    df = pandas.DataFrame(rows)
    df['日期'] = pandas.to_datetime(df['日期'], errors='coerce')
    df['单位净值'] = pandas.to_numeric(df['单位净值'], errors='coerce')
    df['累计净值'] = pandas.to_numeric(df['累计净值'], errors='coerce')
    df['涨跌幅'] = pandas.to_numeric(df['涨跌幅'], errors='coerce')
    return df


@xlwings.func
# 获取基金溢价率、成交量
def get_fund_premium_rate_and_amount(fund_code: str):
    detail = pandas.DataFrame(pysnowball.quote_detail(fund_code))
    row1=detail.loc["quote"][0]
    premium_rate1=row1["premium_rate"]
    amount1 = row1["amount"]
    amount1 = amount1/10000 if amount1 else 0
    return premium_rate1,amount1

@xlwings.func
# 获取xq_a_token
def get_xq_a_token():
    cj = browser_cookie3.load()
    for item in cj:
        if item.name == "xq_a_token" :
            print('%s = %s' % (item.name, item.value))
            str_xq_a_token = 'xq_a_token=' + item.value + ';'
            return str_xq_a_token
    return ';'

@xlwings.func
# 根据排名计算排名分(精确到小数点后2位)
def get_rank_value(length: int, rank: int):
    return round((length-rank)/length*100, 2)


@xlwings.func
def rotate_fund_by_premium_rate_and_20net_asset_value(source_sheets: str, dest_range: str, delay: int):
    '''
    根据20天净值增长率和溢价率来量化轮动基金
    Parameters
    ----------
    source_sheets : Excel里sheet名称
    dest_range : 写回Excel的区域
    delay : 延时
    '''
    xlwings.Book("UniversalRotation.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()  # wb = xlwings.Book(r'UniversalRotation.xlsm')
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    sheet_fund = wb.sheets[source_sheets]
    source_range = 'H2:P' + str(sheet_fund.used_range.last_cell.row)
    print('数据表范围：' + source_range)
    data_fund = pandas.DataFrame(sheet_fund.range(source_range).value,
                                       columns=['基金代码','基金名称','投资种类','20天净值增长率','溢价率','总排名分',
                                                '20天净值增长率排名分','溢价率排名分','成交额(万元)'])
    log_file = open('log-' + source_sheets + str(time.strftime("-%Y-%m-%d-%H-%M-%S", time.localtime())) + '.txt', 'a+')

    for i,fund_code in enumerate(data_fund['基金代码']):
        # 更新溢价率、成交额
        fundPremiumRateValue,fundAmount = get_fund_premium_rate_and_amount(fund_code)
        data_fund.loc[i, '溢价率'] = fundPremiumRateValue
        data_fund.loc[i, '成交额(万元)'] = fundAmount
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
                  + '二十个交易日净值增长率:' + format(str(netAssetValue20Rate), "<10") \
                  + '溢价率:'+ format(str(fundPremiumRateValue), "<10")
        print(log_str)
        print(log_str, file=log_file)

    # 更新'溢价率排名分'
    data_fund = data_fund.sort_values(by='溢价率', ascending=True)
    data_fund.reset_index(drop=True, inplace=True)
    for i,fundPremiumRateValue in enumerate(data_fund['溢价率']):
        # print('溢价率：' + str(fundPremiumRateValue))
        data_fund.loc[i, '溢价率排名分'] = get_rank_value(len(data_fund), i)

    # 更新'20天净值增长率排名分'
    data_fund = data_fund.sort_values(by='20天净值增长率', ascending=False)
    data_fund.reset_index(drop=True, inplace=True)
    for i,netAssetValue in enumerate(data_fund['20天净值增长率']):
        # print('20天净值增长率：' + str(netAssetValue))
        data_fund.loc[i, '20天净值增长率排名分'] = get_rank_value(len(data_fund), i)

    # 更新'总排名分'
    for i,sumValue in enumerate(data_fund['总排名分']):
        data_fund.loc[i, '总排名分'] = data_fund.loc[i, '溢价率排名分'] +\
                                           data_fund.loc[i, '20天净值增长率排名分']
        # print('总排名分'+str(data_fund.loc[i, '总排名分']))
    data_fund = data_fund.sort_values(by='总排名分', ascending=False)
    data_fund.reset_index(drop=True, inplace=True)
    data_fund.index += 1
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()

    # 更新数据到原Excel
    sheet_fund.range(dest_range).value = data_fund
    wb.save()
    os.system("pause")

@xlwings.func
# 轮动20天净值增长和溢价率选LOF、ETF和封基
def rotate_LOF_ETF():
    print("------------------------20天净值增长率和溢价率轮动LOF、ETF和封基-------------------------------------------------")
    rotate_fund_by_premium_rate_and_20net_asset_value('20天净值增长率和溢价率轮动LOF、ETF和封基','G1', 10)


@xlwings.func
# 轮动20天净增和溢价率选债券和境外基金
def rotate_abroad_fund():
    print("-----------------------20天净值增长率和溢价率轮动债券和境外基金---------------------------------------------------")
    rotate_fund_by_premium_rate_and_20net_asset_value('20天净值增长率和溢价率轮动债券和境外基金', 'G1', 10)


@xlwings.func
# 更新可转债实时数据：价格、涨跌幅、转股价、转股价值、溢价率、双低值、到期时间、剩余年限、剩余规模、成交金额、换手率、税前收益、振幅等
def refresh_convertible_bond():
    print("更新可转债实时数据：价格、涨跌幅、转股价、转股价值、溢价率、双低值、到期时间、剩余年限、剩余规模、成交金额、换手率、税前收益、振幅等")
    xlwings.Book("UniversalRotation.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None
    pysnowball.set_token(get_xq_a_token())

    source_sheets = '可转债实时数据'
    sheet_fund = wb.sheets[source_sheets]
    source_range = 'B2:R' + str(sheet_fund.used_range.last_cell.row)
    print('数据表范围：' + source_range)
    data_fund = pandas.DataFrame(sheet_fund.range(source_range).value,
                                       columns=['转债代码','转债名称','当前价','涨跌幅','转股价','转股价值','溢价率','双低值',
                                                '到期时间','剩余年限','剩余规模','成交金额','换手率','税前收益','最高价','最低价','振幅'])
    log_file = open('log-' + source_sheets + str(time.strftime("-%Y-%m-%d-%H-%M-%S", time.localtime())) + '.txt', 'a+')

    for i,fund_code in enumerate(data_fund['转债代码']):
        if str(fund_code).startswith('11') or str(fund_code).startswith('13'):
            fund_code_str = ('SH' + str(fund_code))[0:8]
        elif str(fund_code).startswith('12'):
            fund_code_str = ('SZ' + str(fund_code))[0:8]
        detail = pandas.DataFrame(pysnowball.quote_detail(fund_code_str))
        row1 = detail.loc["quote"][0]
        data_fund.loc[i, '当前价'] = row1["current"]
        data_fund.loc[i, '涨跌幅'] = row1["percent"]
        data_fund.loc[i, '转股价'] = row1["conversion_price"]
        data_fund.loc[i, '转股价值'] = row1["conversion_value"]
        data_fund.loc[i, '溢价率'] = row1["premium_rate"]
        data_fund.loc[i, '双低值'] = row1["current"] + row1["premium_rate"]
        data_fund.loc[i, '到期时间'] = str(time.strftime("%Y-%m-%d", time.localtime(row1["maturity_date"]/1000)))
        data_fund.loc[i, '剩余年限'] = row1["remain_year"]
        data_fund.loc[i, '剩余规模'] = row1["outstanding_amt"] / 100000000 if row1["outstanding_amt"] else 0
        data_fund.loc[i, '成交金额'] = row1["amount"] / 10000 if row1["amount"] else 0
        data_fund.loc[i, '换手率'] = (data_fund.loc[i, '成交金额'] / 10000 / row1["current"]) / (data_fund.loc[i, '剩余规模'] / 100)
        data_fund.loc[i, '税前收益'] = row1["benefit_before_tax"]
        data_fund.loc[i, '最高价'] = row1["high"]
        data_fund.loc[i, '最低价'] = row1["low"]
        if row1["high"] and row1["low"]:
            data_fund.loc[i, '振幅'] = (row1["high"] - row1["low"]) / row1["low"]
        else:
            data_fund.loc[i, '振幅'] = '停牌'
        log_str = 'No.' + format(str(i), "<6") + format(str(fund_code_str), "<10") \
                  + format(data_fund.loc[i, '转债名称'], "<15") \
                  + '当前价:' + format(str(row1["current"]), "<10") \
                  + '溢价率:' + format(str(row1["premium_rate"]), "<10") \
                  + '涨跌幅:'+ format(str(row1["percent"]), "<10")
        print(log_str)
        print(log_str, file=log_file)

    data_fund = data_fund.sort_values(by='溢价率')
    data_fund.reset_index(drop=True, inplace=True)
    data_fund.index +=1
    print(data_fund)
    print(data_fund, file=log_file)

    log_file.close()
    # 更新原Excel
    sheet_fund.range('A1').value = data_fund
    wb.save()
    os.system("pause")

@xlwings.func
# 更新低溢价可转债数据
def refresh_premium_rate_convertible_bond():
    print("--------------------------------------更新低溢价可转债数据----------------------------------------------------")
    xlwings.Book("UniversalRotation.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['可转债实时数据']
    source_range = 'B2:R' + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['转债代码','转债名称','当前价','涨跌幅','转股价','转股价值','溢价率','双低值',
                                                '到期时间','剩余年限','剩余规模','成交金额','换手率','税前收益','最高价','最低价','振幅'])
    data_fund_destination = data_fund_source[['转债代码','转债名称','当前价','溢价率','剩余规模']]
    data_fund_destination = data_fund_destination.sort_values(by='溢价率')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    # 更新低溢价可转债轮动sheet
    sheet_dest = wb.sheets['低溢价可转债轮动']
    sheet_dest.range('H2').value = data_fund_destination
    wb.save()
    os.system("pause")


@xlwings.func
# 更新双低可转债数据
def refresh_price_and_premium_rate_convertible_bond():
    print("----------------------------------------更新双低可转债数据----------------------------------------------------")
    xlwings.Book("UniversalRotation.xlsm").set_mock_caller()
    wb = xlwings.Book.caller()
    pandas.options.display.max_columns = None
    pandas.options.display.max_rows = None

    sheet_src = wb.sheets['可转债实时数据']
    source_range = 'B2:R' + str(sheet_src.used_range.last_cell.row)
    data_fund_source = pandas.DataFrame(sheet_src.range(source_range).value,
                                       columns=['转债代码','转债名称','当前价','涨跌幅','转股价','转股价值','溢价率','双低值',
                                                '到期时间','剩余年限','剩余规模','成交金额','换手率','税前收益','最高价','最低价','振幅'])
    data_fund_destination = data_fund_source[['转债代码','转债名称','当前价','溢价率','双低值','剩余规模']]
    data_fund_destination = data_fund_destination.sort_values(by='双低值')
    data_fund_destination.reset_index(drop=True, inplace=True)
    data_fund_destination.index +=1
    print(data_fund_destination)

    # 更新双低可转债轮动sheet
    sheet_dest = wb.sheets['双低可转债轮动']
    sheet_dest.range('H2').value = data_fund_destination
    wb.save()
    os.system("pause")

def main():
    #删除旧log
    for eachfile in os.listdir('./'):
        filename = os.path.join('./', eachfile)
        if os.path.isfile(filename) and filename.startswith("./log") :
            os.remove(filename)

    rotate_LOF_ETF()
    rotate_abroad_fund()
    refresh_convertible_bond()
    refresh_premium_rate_convertible_bond()
    refresh_price_and_premium_rate_convertible_bond()

if __name__ == "__main__":
    main()