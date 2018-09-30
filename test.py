# import matplotlib.pyplot as plt
# from matplotlib.font_manager import FontProperties
# font=FontProperties(fname=r'c:\windows\fonts\SimHei.ttf')
# plt.clf()  # 清空画布
# plt.plot([-5, 2, 3], [-4, 5, 6])
# plt.xlabel(u"横轴",fontproperties=font)
# plt.ylabel(u"纵轴",fontproperties=font)
# plt.title("pythoner.com")
# plt.show()


import matplotlib.pyplot as plt
import pickle
from matplotlib.font_manager import FontProperties
font=FontProperties(fname=r'c:\windows\fonts\SimHei.ttf')
# ctsj=invoker.wsd("501015.SH", "close,volume,amt,chg,pct_chg,turn", "2018-01-01", "2018-09-18", "")
# ctsj.index=pd.to_datetime(ctsj.index)
#保存
input=open(r'd:\ctsj.pkl','rb')
ctsj=pickle.load(input)
ctsj.to_excel(r'd:\ctsj.xlsx')
ctsj['2018-05-25':'2018-08-26'].AMT.mean()
ctsj['2018-08-27':'2018-08-27'].AMT
plt.plot(ctsj['2018-05-27':'2018-08-27'].AMT)
plt.title('财通升级501015最近3个月日成交额',fontproperties=font,fontsize=16)
plt.show()

((ctsj.NAV/ctsj.CLOSE-1)*100).plot()