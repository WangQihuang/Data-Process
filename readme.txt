Agilent ICPMS8900仪器数据处理脚本

主要功能：
快速筛去仪器导出数据的无用行列
根据设定的平行样品个数快速计算样品浓度的平均值和标准差


python的安装：
文件夹中python-3.8.3-amd64.exe为pyton安装包
双击安装包
勾选Install launcher for all users和Add python 3.8 to PATH两个选项
选择Customize installation进行自定义安装
勾选Optinal Features下的所有选项
勾选Advanced Features下的所有选项（最后两个可不选）
安装地址默认的即可，更改自定义地址不要有中文字符！
点击Install开始安装

python环境配置：
win+R输入cmd进入命令行窗口
检查已经安装的python第三方包
pip list
安装程序运行所必须的python第三方包
pip install numpy
pip install openpyxl
pip install pandas
pip install pip install xlrd==1.2.0
如果安装了xlrd 1.2.0以上版本，需卸载后安装旧版！！
pip install pip uninstall xlrd

CSVdataprocess.py
为通用版本，按照命令行窗口提示的内容按顺序输入相关参数即可。

CSVdataprocess@username.py
为魔改版本，需将outputfilepass中UserName替换为当前电脑用户名方可使用
注意：文件路径中不可有中文！！

文件夹中Quicker.x64.1.24.34.0.msi为Quicker安装包
内置脚本可快速提取文件路径（搜索：取路径）
脚本url：https://getquicker.net/sharedaction?code=d645435d-1319-4ec1-a451-08d787848547

