### (一)、使用Python+Excel可视化你的持仓和最新排名的差异，解决买什么卖什么的烦恼
1. 这里以"低溢价可转债策略"、"双低可转债策略"为例子，我做了2个sheet【低溢价可转债轮动】和【双低可转债轮动】。
2. 从券商软件下载最新的低溢价可转债持仓和双低可转债持仓的Excel表格，并分别更新到“我的低溢价可转债持仓”、“我的双低可转债持仓”。
3. 更新sheet页里面的”最新低溢价可转债排名“、”最新双低可转债排名“、其他任何自定义的”最新VIP可转债排名“，即可看到轮动结果了。

注1：这里依赖的是Excel自带的各种公式，忽然发现Excel解放手脚的强大了吧？神不神奇惊不惊喜，意不意外?

注2：每个sheet页我都做了双排名（即自定义的”最新VIP可转债排名“），满足用户自定义的需求。

注3：这里的2个sheet页仅是2个可转债的例子，填充了持仓和最新排名。用户可以将它应用于任何的“依据各种因子进行排名的量化轮动策略”。
大家可以将这两个sheet作为壳子，替换成任何你们自己的轮动品种，比如小市值低价股票策略、低估股票策略、A/H轮动等等。


### (二)、关于《20天净值增长率和溢价率轮动LOF、ETF和封基》和《20天净值增长率和溢价率轮动债券和境外基金》
1. 安装Python3：https://www.python.org/ftp/python/3.8.7/python-3.8.7-amd64.exe
2. 打开cmd窗口输入：pip install xlwings pandas requests pysnowball
3. 启用excel中的xlwings宏：
   (a)、命令行安装加载项：xlwings addin install。
   (b)、在excel中启用加载项： 文件>选项>信任中心>信任中心设置>宏设置 中，选择“启用所有并勾选”并勾选“对VBA对象模型的信任访问”。 
4. 点击“更新LOF/ETF/封基策略”、“更新债券/境外策略”按钮，即可更新这2个策略的最新的排名数据。

注4：Python调用API获取溢价率前需要设置token，有20天有效期，可以参考https://blog.crackcreed.com/diy-xue-qiu-app-shu-ju-api/来获取token，然后修改UniversalRotation.py里的下面这段code里的xq_a_token即可：
pysnowball.set_token('xq_a_token=e8119f7d7a050cdbfa822fa0da4de5bec1ee0dc7;')

注5：作为一个Android程序员，从2022.2月份开始边学边练第一次写Python，语法格式肯定不完美，勿喷，仅仅是为了解放调仓的苦恼而写的小玩意，分享出来仅用于学习研究，不可用于商业用途！
