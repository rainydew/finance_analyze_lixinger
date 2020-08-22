# coding: utf-8
from typing import List
import requests
import json
from easygui import msgbox

stocks = {}
token = input("请输入token ")
start_year = input("请输入起始年份(如2014) ")
end_year = input("请输入结束年份(如2019，理杏仁限制10年区间) ")
while True:
    code = input("请输入股票代码，如果要结束输入，请直接回车 ")
    if not code:
        break
    name = input(f"请输入代码{code}对应的股票名称 ")
    stocks[code] = name

csv_file = ""
csv_list = []
parse_dict = {'ps.toi': '营业总收入', 'ps.oi': '营业收入', 'ps.oc': '营业成本', 'ps.tas': '税金及附加', 'ps.se': '销售费用', 'ps.ae': '管理费用', 'ps.rade': '研发费用', 'ps.fe': '财务费用', 'ps.oic': '其他收益', 'ps.ivi': '投资收益', 'ps.ciofv': '公允价值变动收益', 'ps.cilor': '信用减值损失', 'ps.ailor': '资产减值损失', 'ps.adi': '资产处置收益', 'ps.noi': '营业外收入', 'ps.tp': '利润总额', 'ps.np_s_r': '净利润率', 'bs.ta': '资产合计', 'bs.fa': '固定资产', 'bs.cip': '在建工程', 'bs.lwi_ta_r': '有息负债率', 'bs.tl_ta_r': '资产负债率', 'bs.tncl': '非流动负债合计', 'bs.toe': '股东权益合计', 'cfs.crfscapls': '销售商品、提供劳务收到的现金', 'cfs.ncffoa': '经营活动产生的现金流量净额', 'cfs.ncffia': '投资活动产生的现金流量净额', 'cfs.ncfffa': '筹资活动产生的现金流量净额', 'm.ta_to': '资产周转率', 'm.gp_r': '毛利率(GM)', 'm.i_ds': '存货周转天数', 'm.ar_ds': '应收账款周转天数', 'm.c_r': '流动比率', 'm.q_r': '速动比率'}
year_list = list(range(int(start_year), int(end_year) + 1))

ly = len(year_list)
cols = ["公司", "收入质量", *["营业收入增长率"] * ly, *["营业成本率"] * ly, *["营业费用率"] * ly, *["销售商品、提供劳务收到的现金/营业收入比值"] * ly,
        *["毛利率"] * ly, *["核心利润率"] * ly, *["净利润率"] * ly, *["资产负债率"] * ly, *["有息负债率"] * ly, *["长期资金占不动产及设备比率"] * ly,
        *["流动比率"] * ly, *["速动比率"] * ly, *["有形资产占比"] * ly, *["利润总额占有形资产比例"] * ly, *["资产周转率"] * ly,
        *["存货周转天数"] * ly, *["应收账款周转天数"] * ly, *["经营活动产生的现金流量净额", "投资活动产生的现金流量净额", "筹资活动产生的现金流量净额"] * ly]
csv_list.append(cols)
sec_r = ["年份", "last"]
cir = (len(cols) - 2) / ly - 3
[sec_r.extend(year_list) for _ in range(int(cir))]
[sec_r.extend([i, i, i]) for i in year_list]
csv_list.append(sec_r)
print(csv_list)


for stock, stock_name in stocks.items():
    row = [stock_name]
    param = {
        "token": token,
        "startDate": f"{start_year}-12-31",
        "endDate": f"{end_year}-12-31",
        "stockCodes": [
            stock
        ],
        "metricsList": [
            "y.ps.toi.t",
            "y.ps.toi.t_y2y",
            "y.ps.oi.t",
            "y.ps.oi.t_y2y",
            "y.ps.oc.t",
            "y.ps.oc.t_y2y",
            "y.ps.tas.t",
            "y.ps.tas.t_y2y",
            "y.ps.se.t",
            "y.ps.se.t_y2y",
            "y.ps.ae.t",
            "y.ps.ae.t_y2y",
            "y.ps.rade.t",
            "y.ps.rade.t_y2y",
            "y.ps.fe.t",
            "y.ps.fe.t_y2y",
            "y.ps.oic.t",
            "y.ps.oic.t_y2y",
            "y.ps.ivi.t",
            "y.ps.ivi.t_y2y",
            "y.ps.ciofv.t",
            "y.ps.ciofv.t_y2y",
            "y.ps.cilor.t",
            "y.ps.cilor.t_y2y",
            "y.ps.ailor.t",
            "y.ps.ailor.t_y2y",
            "y.ps.adi.t",
            "y.ps.adi.t_y2y",
            "y.ps.noi.t",
            "y.ps.noi.t_y2y",
            "y.ps.tp.t",
            "y.ps.tp.t_y2y",
            "y.ps.np_s_r.t",
            "y.ps.np_s_r.t_y2y",
            "y.bs.ta.t",
            "y.bs.ta.t_y2y",
            "y.bs.fa.t",
            "y.bs.fa.t_y2y",
            "y.bs.cip.t",
            "y.bs.cip.t_y2y",
            "y.bs.lwi_ta_r.t",
            "y.bs.lwi_ta_r.t_y2y",
            "y.bs.tl_ta_r.t",
            "y.bs.tl_ta_r.t_y2y",
            "y.bs.tncl.t",
            "y.bs.tncl.t_y2y",
            "y.bs.toe.t",
            "y.bs.toe.t_y2y",
            "y.cfs.crfscapls.t",
            "y.cfs.crfscapls.t_y2y",
            "y.cfs.ncffoa.t",
            "y.cfs.ncffoa.t_y2y",
            "y.cfs.ncffia.t",
            "y.cfs.ncffia.t_y2y",
            "y.cfs.ncfffa.t",
            "y.cfs.ncfffa.t_y2y",
            "y.m.ta_to.t",
            "y.m.ta_to.t_y2y",
            "y.m.gp_r.t",
            "y.m.gp_r.t_y2y",
            "y.m.i_ds.t",
            "y.m.i_ds.t_y2y",
            "y.m.ar_ds.t",
            "y.m.ar_ds.t_y2y",
            "y.m.c_r.t",
            "y.m.c_r.t_y2y",
            "y.m.q_r.t",
            "y.m.q_r.t_y2y",
        ]
    }


    res = requests.post("https://open.lixinger.com/api/a/stock/fs/non_financial", json=param)
    res = res.json()
    print(json.dumps(res, indent=2, ensure_ascii=False))

    info = res["data"]
    info = [x for x in info if "-12-31T" in x["standardDate"]]
    res_dict = {}
    res_rate_dict = {}

    for x in info:
        year = int(x["standardDate"][:4])
        for d_type, d in x["y"].items():
            for p_type, p in d.items():
                arg = parse_dict[f"{d_type}.{p_type}"]
                print(year, arg, p["t"], str(int(p["t_y2y"] * 10000 + 0.5) / 100) + "%" if p.get("t_y2y") else "NA")
                if arg in res_dict:
                    res_dict[arg][year_list.index(year)] = p["t"]
                    res_rate_dict[arg][year_list.index(year)] = int(p["t_y2y"] * 10000 + 0.5) / 100 if p.get(
                        "t_y2y") else None
                else:
                    res_rate_dict[arg]: List[float] = [None] * len(year_list)
                    res_dict[arg]: List[float] = [None] * len(year_list)
                    res_dict[arg][year_list.index(year)] = p["t"]
                    res_rate_dict[arg][year_list.index(year)] = int(p["t_y2y"] * 10000 + 0.5) / 100 if p.get(
                        "t_y2y") else None

    print(res_dict)
    print(res_rate_dict)
    yingyeshouru_last = res_dict["营业收入"][-1]
    yingyeshouru = res_dict["营业收入"]
    qita_yingyewaishouru = res_dict["其他收益"][-1] + res_dict["投资收益"][-1] + res_dict["公允价值变动收益"][-1] + res_dict["信用减值损失"][
        -1] - res_dict["资产减值损失"][-1] + res_dict["资产处置收益"][-1] + res_dict["营业外收入"][-1]
    shourubili = yingyeshouru_last / (yingyeshouru_last + qita_yingyewaishouru)
    print("收入分析")
    print("收入质量", shourubili, "收入质量好，营收占比大于90%" if shourubili >= 0.9 else "收入质量不好，营收占比小于90%")
    row.append(shourubili)

    yingyeshouruzengzhanglv = {year_list[i]: str(x) + "%" for i, x in enumerate(res_rate_dict["营业收入"])}
    print("营业收入增长率", yingyeshouruzengzhanglv)
    row.extend(yingyeshouruzengzhanglv.values())

    print("成本分析")
    yingyechengben = res_dict["营业成本"]
    yingyechengbenlv = [yingyechengben[i] / x for i, x in enumerate(yingyeshouru)]
    yingyechengbenlv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(yingyechengbenlv)}
    print("营业成本率", yingyechengbenlv_d)
    row.extend(yingyechengbenlv_d.values())

    yingyefeiyonglv = [0] * len(year_list)
    for i in range(len(year_list)):
        yingyefeiyonglv[i] = (res_dict["销售费用"][i] or 0 + res_dict["研发费用"][i] or 0 + res_dict["管理费用"][i] or 0 +
                              res_dict["财务费用"][i] or 0) / yingyeshouru[i]
    yingyefeiyonglv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(yingyefeiyonglv)}
    print("营业费用率", yingyefeiyonglv_d)
    row.extend(yingyefeiyonglv_d.values())

    print("随着营业收入的增长，成本率，费用率保持稳定或逐渐小幅下降，在此基础上，如果比同行更低，说明成本控制越好")
    print("利润分析")
    xiaoshou_laowuxianjin = res_dict["销售商品、提供劳务收到的现金"]
    lirunzhenshixing = [xiaoshou_laowuxianjin[i] / x for i, x in enumerate(yingyeshouru)]
    print("销售商品、提供劳务收到的现金/营业收入比值", lirunzhenshixing)
    row.extend(lirunzhenshixing)

    maolilv = res_dict["毛利率(GM)"]
    maolilv = {year_list[i]: x for i, x in enumerate(maolilv)}
    print("毛利率(GM)", maolilv)
    row.extend(maolilv.values())

    hexinlirunlv = [0] * len(year_list)
    shuijinfujia = res_dict["税金及附加"]
    yingyezongshouru = res_dict["营业总收入"]
    for i in range(len(year_list)):
        hexinlirunlv[i] = (yingyeshouru[i] - yingyechengben[i] - (
                    res_dict["销售费用"][i] or 0 + res_dict["研发费用"][i] or 0 + res_dict["管理费用"][i] or 0 + res_dict["财务费用"][i] or 0) - shuijinfujia[i]) / yingyezongshouru[i]
    hexinlirunlv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(hexinlirunlv)}
    print("核心利润率", hexinlirunlv_d)
    row.extend(hexinlirunlv_d.values())

    jinglirunlv = res_dict["净利润率"]
    jinglirunlv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(jinglirunlv)}
    print("净利润率", jinglirunlv_d)
    row.extend(jinglirunlv_d.values())

    print("财务结构")
    zichanfuzhailv = res_dict["资产负债率"]
    zichanfuzhailv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(zichanfuzhailv)}
    print("资产负债率", zichanfuzhailv_d)
    row.extend(zichanfuzhailv_d.values())

    print("经营稳定的公司，资产负债率一般不超过50-60%，金融地产类除外")
    youxifuzhailv = res_dict["有息负债率"]
    youxifuzhailv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(youxifuzhailv)}
    print("有息负债率", youxifuzhailv_d)
    row.extend(youxifuzhailv_d.values())

    print("一家公司越是强势，在行业内的地位和话语权越高，有息负债率越低")
    changqizijinzhanbi = [0] * len(year_list)
    gudongquanyiheji = res_dict["股东权益合计"]
    feiliudongfuzhaiheji = res_dict["非流动负债合计"]
    gudingzichan = res_dict["固定资产"]
    zaijiangongcheng = res_dict["在建工程"]
    for i in range(len(year_list)):
        changqizijinzhanbi[i] = (gudongquanyiheji[i] + feiliudongfuzhaiheji[i]) / (
                    gudingzichan[i] + zaijiangongcheng[i])
    changqizijinzhanbi_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(changqizijinzhanbi)}
    print("长期资金占不动产及设备比率", changqizijinzhanbi_d)
    row.extend(changqizijinzhanbi_d.values())

    print("长期资产占不动产及设备比率大于100%，则说明企业的长期资金充足")
    print("偿债能力")
    liudongbilv = res_dict["流动比率"]
    liudongbilv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(liudongbilv)}
    print("流动比率", liudongbilv_d)
    row.extend(liudongbilv_d.values())

    print("""流动比率在1~2之间，算是尚可接受的状态。
    如果流动比率小于1，是个比较糟糕的状态
    流动比率远大于2，说明企业有较多的资金滞留在流动资产上未加以更好运用""")
    sudongbilv = res_dict["速动比率"]
    sudongbilv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(sudongbilv)}
    print("速动比率", sudongbilv_d)
    row.extend(sudongbilv_d.values())

    print("""速动比率在1左右，说明扣掉了存货和预付账款，公司剩下的流动资产依旧比流动负债多。
    速动比率在0.5~1之间，快速变现的资产只比流动负债略少一些。
    速动比率小于0.5，偿债能力较差
    速动比率远大于1，同样可能存在大量资金闲置的问题""")
    print("资产结构")
    zichanheji = res_dict["资产合计"]
    youxingzichanbili = [0] * len(year_list)
    for i in range(len(year_list)):
        youxingzichanbili[i] = (gudingzichan[i] + zaijiangongcheng[i]) / zichanheji[i]
    youxingzichanbili_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(youxingzichanbili)}
    print("有形资产占比", youxingzichanbili_d)
    row.extend(youxingzichanbili_d.values())

    print("有形资产占比在30%以上的公司，被归为重资产公司，小于30%为轻资产公司")
    lirunzongezhanyouxingzichanbili = [0] * len(year_list)
    lirunzonge = res_dict["利润总额"]
    for i in range(len(year_list)):
        lirunzongezhanyouxingzichanbili[i] = lirunzonge[i] / (gudingzichan[i] + zaijiangongcheng[i])
    lirunzongezhanyouxingzichanbili_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(lirunzongezhanyouxingzichanbili)}
    print("利润总额占有形资产比例", lirunzongezhanyouxingzichanbili_d)
    row.extend(lirunzongezhanyouxingzichanbili_d.values())

    print("重资产公司的比值大于10%，就可以认为公司经营状况不错、赚钱能力强，具有进一步考察的价值(轻资产意义不大)")
    print("运营能力")
    zichanzhouzhuanlv = res_dict["资产周转率"]
    zichanzhouzhuanlv_d = {year_list[i]: str(int(x * 10000 + 0.5) / 100) + "%" for i, x in enumerate(zichanzhouzhuanlv)}
    print("资产周转率", zichanzhouzhuanlv_d)
    row.extend(zichanzhouzhuanlv_d.values())

    print("总资产周转率越高，说明公司投入总资产赚到的钱越多")
    cunhuozhouzhuantianshu = res_dict["存货周转天数"]
    cunhuozhouzhuantianshu_d = {year_list[i]: int(x * 10 + 0.5) / 10 for i, x in enumerate(cunhuozhouzhuantianshu)}
    print("存货周转天数", cunhuozhouzhuantianshu_d)
    row.extend(cunhuozhouzhuantianshu_d.values())

    print("存货周转天数越短，代表产品占用资金的时间越短，公司对存货的运营能力越高")
    yingshouzhangkuanzhouzhuantianshu = res_dict["应收账款周转天数"]
    yingshouzhangkuanzhouzhuantianshu_d = {year_list[i]: int(x * 10 + 0.5) / 10 for i, x in enumerate(yingshouzhangkuanzhouzhuantianshu)}
    print("应收账款周转天数", yingshouzhangkuanzhouzhuantianshu_d)
    row.extend(yingshouzhangkuanzhouzhuantianshu_d.values())

    print("应收账款周转天数越短，代表收回赊销货款的速度越快，公司的收现能力越高")
    print("现金流量分析")
    jingyingxianjinliu = res_dict["经营活动产生的现金流量净额"]
    touzixianjinliu = res_dict["投资活动产生的现金流量净额"]
    chouzixianjinliu = res_dict["筹资活动产生的现金流量净额"]
    print("  年份  ", year_list)
    print("经营活动现金流量净额", jingyingxianjinliu)
    print("投资活动现金流量净额", touzixianjinliu)
    print("筹资活动现金流量净额", chouzixianjinliu)
    row.extend(jingyingxianjinliu)
    row.extend(touzixianjinliu)
    row.extend(chouzixianjinliu)

    csv_list.append(row)

print("csv_list", csv_list)


def add_col(*arg):
    global csv_file
    for a in arg:
        csv_file += f"{a},"

for j in range(0, 2 + len(stocks)):
    add_col(csv_list[j][0], csv_list[j][1], csv_list[j][2])
    csv_file += "\n"
csv_file += "\n"


for i in range(2, len(cols) - ly * 3, ly):
    for j in range(0, 2 + len(stocks)):
        add_col(csv_list[j][0])
        add_col(*csv_list[j][i:i+ly]) if j else add_col(csv_list[j][i])
        csv_file += "\n"
    csv_file += "\n"

j = 2
t = 0
tl = ["经营活动产生的现金流量净额", "投资活动产生的现金流量净额", "筹资活动产生的现金流量净额"]
add_col("年份", *year_list)
csv_file += "\n"
for i in range(len(cols) - ly * 3, len(cols), ly):
    add_col(tl[t])
    add_col(*csv_list[j][i:i+ly])
    csv_file += "\n"
    t += 1

while True:
    try:
        with open("temp.csv", "w") as f:
            f.write(csv_file)
    except PermissionError:
        msgbox("文件被占用，请关闭打开temp.csv的程序并重试")
    else:
        break
