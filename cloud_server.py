# -*- coding: utf-8 -*-
"""裕安走访系统 - 云端版（内嵌数据，独立运行）
部署于云服务器，监听80端口。数据文件: /opt/visit/visit_results.json
"""
import json, os, io, threading, datetime, urllib.parse
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
RESULTS = os.path.join(BASE, 'visit_results.json')
LOCK = threading.Lock()

DATA = json.loads(r"""{"recs":[{"n":"深圳市艾尼亚电子科技有限公司","assign":"非裕安","tel":"0755-86381649","legal":"赵喜平","grid":"(未匹配)","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新桥街道上寮社区黄埔路130号厂房二层","ind":"批发业","size":"M(中型)","staff":24,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市金蓝建科技有限公司","assign":"非裕安","tel":"18129821276","legal":"罗雪","grid":"鸿都工业园","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区航城街道黄麻布社区簕竹角鸿都工业园3号厂房2层","ind":"批发业","size":"XS(微型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市亚达兴业科技有限公司","assign":"非裕安","tel":"18938652046","legal":"陈忠春","grid":"鸿都工业园","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区航城街道黄麻布社区簕竹角鸿都工业园3号厂房4层","ind":"计算机、通信和其他电子设","size":"S(小型)","staff":8,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳迈联电子有限公司","assign":"非裕安","tel":"13322980799","legal":"刘亚秀","grid":"鸿都工业园","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区航城街道簕竹角社区簕竹角鸿都工业园5号厂房2层","ind":"计算机、通信和其他电子设","size":"XS(微型)","staff":9,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳谜点科技有限公司","assign":"非裕安","tel":"13925661123","legal":"廖海滨","grid":"中心大厦","grade":"A","src":"跨境","risk":"轻涉诉","addr":"深圳市宝安区新桥街道新二社区蚝乡路同方中心大厦1705.1706","ind":"科技推广和应用服务业","size":"S(小型)","staff":12,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳优博胜电子有限公司","assign":"非裕安","tel":"13352908573","legal":"谢君","grid":"中心大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新桥街道新二社区中心路与蚝乡路交汇处同方中心大厦1409-1410","ind":"-","size":"XS(微型)","staff":14,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市安迪生照明有限公司","assign":"黄锋","tel":"18822885445","legal":"班卓娅","grid":"方华工业区","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道大浪社区新安三路一巷10号方华工业区厂房1栋309","ind":"批发业","size":"S(小型)","staff":11,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市益凯尔电子有限公司","assign":"黄锋","tel":"13613003500","legal":"黄美燕","grid":"华利工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区裕安二路404号华利工业区厂房B201、B401、B501、B601","ind":"计算机、通信和其他电子设","size":"S(小型)","staff":136,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市精灵体智能家居产品有限公司","assign":"黄锋","tel":"13480263725","legal":"华椿萱","grid":"建达工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区新安三路118号建达工业区厂房1号202-203","ind":"零售业","size":"S(小型)","staff":14,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市锐途电子商务有限公司","assign":"黄锋","tel":"13510131731","legal":"樊家平","grid":"建达工业区","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道大浪社区新安三路118号建达工业区厂房3号（1-5层）304","ind":"邮政业","size":"XS(微型)","staff":15,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、SDWAN(多仓多网点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市凯凌飞科技有限公司","assign":"黄锋","tel":"15815597512","legal":"樊波东","grid":"建达工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区新安三路118号建达工业区厂房3号（1-5层）401","ind":"零售业","size":"XS(微型)","staff":8,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市赛杰电子有限公司","assign":"黄锋","tel":"13554737350","legal":"覃新刚","grid":"新柯城科技工业园","grade":"A","src":"跨境","risk":"轻涉诉","addr":"深圳市宝安区新安街道28区大宝路51号新柯城工业园1栋3楼西面","ind":"文教、工美、体育和娱乐用","size":"S(小型)","staff":21,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市陶知之文化传播有限公司","assign":"江静竹","tel":"18998730909","legal":"陈桂红","grid":"(未匹配)","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区前进一路92号玖悦雅轩434","ind":"零售业","size":"S(小型)","staff":8,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市嘉明欣科技有限公司","assign":"江静竹","tel":"18926757483","legal":"张仲","grid":"成悦里","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道凤凰岗社区中熙香缤山花园7栋3102","ind":"批发业","size":"XS(微型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市宏科特电子科技有限公司","assign":"江静竹","tel":"13592950563","legal":"李永华","grid":"成悦里","grade":"A","src":"他网","risk":"需关注","addr":"深圳市宝安区西乡街道凤凰岗社区中熙香缤山花园8栋1108、1111-1114","ind":"电子元件制造","size":"S(小型)","staff":35,"stock":"否(他网客户)","rec":"【核心】自有服务器在联通机房→迁电信IDC/天翼云+专线+固定IP一揽子；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有","domain":"hktelec.com","op":"联通","dns":"A记录(直连)","deploy":"中国联通"},{"n":"深圳市华必达科技有限公司","assign":"江静竹","tel":"0755-81773879","legal":"李鹏","grid":"成悦里","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道凤凰岗社区中熙香缤山花园9栋成悦里408","ind":"批发业","size":"XS(微型)","staff":11,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳绿量能源电池有限公司","assign":"江静竹","tel":"13651484135","legal":"殷开红","grid":"福城时代广场","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区13区宝民一路221号福城时代广场一单元1912","ind":"零售业","size":"XS(微型)","staff":8,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市资合商贸有限公司","assign":"江静竹","tel":"16675158822","legal":"卢洁娜","grid":"福城时代广场","grade":"C","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区13区宝民一路221号福城时代广场一单元2112","ind":"批发业","size":"S(小型)","staff":20,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市扬帆纸制品有限公司","assign":"江静竹","tel":"18676733566","legal":"陈长宝","grid":"宏发玺玥大厦","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道大浪社区27区创业二路111号宏发玺玥大厦3502A","ind":"科技推广和应用服务业","size":"S(小型)","staff":15,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市玖誉纸制品有限公司","assign":"江静竹","tel":"18676733566","legal":"陈长宝","grid":"宏发玺玥大厦","grade":"A","src":"跨境","risk":"轻涉诉","addr":"深圳市宝安区新安街道大浪社区27区创业二路111号宏发玺玥大厦3502A","ind":"批发业","size":"S(小型)","staff":11,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市佩特拉实业有限公司","assign":"江静竹","tel":"17665257751","legal":"陈长宝","grid":"宏发玺玥大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区27区创业二路111号宏发玺玥大厦3502D","ind":"批发业","size":"S(小型)","staff":12,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"万事威科技（深圳）有限公司","assign":"江静竹","tel":"0755-27805686","legal":"张泽铭","grid":"宏发玺玥大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区27区创业二路111号宏发玺玥大厦609","ind":"零售业","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市拓刻科技有限公司","assign":"江静竹","tel":"13631697212","legal":"陈敏玲","grid":"众里创业社区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区大宝路49号众里创业社区洪浪北204","ind":"零售业","size":"S(小型)","staff":6,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市加雪龙科技有限公司","assign":"江静竹","tel":"0755-23322700","legal":"陈祥宝","grid":"众里创业社区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道大浪社区大宝路49号众里创业社区洪浪北E3","ind":"批发业","size":"S(小型)","staff":10,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市科及之星电脑有限公司","assign":"刘成芳","tel":"13724325755","legal":"陈财明","grid":"(未匹配)","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区新安四路263号旭仕达名苑215","ind":"软件和信息技术服务业","size":"XS(微型)","staff":6,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市菲弘特贸易有限公司","assign":"刘成芳","tel":"13925226919","legal":"胡教洪","grid":"城西工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道河西社区城西工业区2栋501","ind":"批发业","size":"XS(微型)","staff":8,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市明优实业有限公司","assign":"刘成芳","tel":"13699896158","legal":"刘树中","grid":"城西工业区","grade":"A","src":"他网","risk":"轻涉诉","addr":"深圳市宝安区西乡街道河西社区城西工业区5栋2层","ind":"清洁服务","size":"S(小型)","staff":33,"stock":"否(他网客户)","rec":"【核心】自有服务器在联通机房→迁电信IDC/天翼云+专线+固定IP一揽子；商企宽带+组网WiFi整体策反","domain":"mingyousz.com","op":"联通","dns":"A记录(直连)","deploy":"中国联通"},{"n":"深圳科时达表业有限公司","assign":"刘成芳","tel":"13923763323","legal":"潘铠勉","grid":"凤凰岗第三工业区","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道凤凰岗第三工业区B7栋5楼","ind":"仪器仪表制造业","size":"S(小型)","staff":31,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市春秋电子科技有限公司","assign":"刘成芳","tel":"18926774811","legal":"郭官苑","grid":"凤凰岗第三工业区","grade":"A","src":"跨境","risk":"需关注","addr":"深圳市宝安区西乡街道凤凰岗社区凤凰岗第三工业区B2401、B2501","ind":"科技推广和应用服务业","size":"S(小型)","staff":33,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市凝炼科技有限公司","assign":"刘成芳","tel":"13632599384","legal":"傅廷宝","grid":"凤凰岗第三工业区","grade":"B","src":"他网","risk":"未扫描","addr":"深圳市宝安区西乡街道铁岗社区桃花源科技创新生态园B5栋3A","ind":"半导体","size":"XS(微型)","staff":5,"stock":"否(他网客户)","rec":"【核心】自有服务器在联通机房→迁电信IDC/天翼云+专线+固定IP一揽子；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有","domain":"nlkeji.com","op":"联通","dns":"A记录(直连)","deploy":"国内IDC(中国联通线路)"},{"n":"深圳市格天光电有限公司","assign":"刘成芳","tel":"0755-29656860","legal":"王志成","grid":"凤凰岗第一工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道凤凰岗社区凤凰岗第一工业区厂房B座301","ind":"电气机械和器材制造业","size":"S(小型)","staff":45,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳中瀚通信有限公司","assign":"刘成芳","tel":"18718519000","legal":"刘时青","grid":"凤凰岗第一工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道凤凰岗社区凤凰岗路17号705","ind":"零售业","size":"XS(微型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳柒颜电子商务有限公司","assign":"刘成芳","tel":"13825209335","legal":"黎晓勤","grid":"富强大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道流塘社区62区麻布村1栋1-3富强大厦1201","ind":"批发业","size":"XS(微型)","staff":9,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳菲尔珀科技有限公司","assign":"刘成芳","tel":"15250680555","legal":"黎晓勤","grid":"富强大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道流塘社区62区麻布村1栋1-3富强大厦1201","ind":"批发业","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市鼎创中正科技有限公司","assign":"刘成芳","tel":"18601658083","legal":"曹东进","grid":"富强大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道流塘社区62区麻布村1栋1-3富强大厦506","ind":"批发业","size":"S(小型)","staff":9,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市众锦鑫科技有限公司","assign":"刘成芳","tel":"13689516366","legal":"邱杏允","grid":"丽景食城","grade":"C","src":"他网","risk":"轻涉诉","addr":"深圳市宝安区西乡街道富华社区前进二路丽景城4栋二单元305","ind":"软件开发","size":"XS(微型)","staff":10,"stock":"否(他网客户)","rec":"【核心】CNAME建站托管→天翼云建站迁移轻量切入；SDWAN+云电脑；商企宽带+组网WiFi整体策反；跨境专线(有进出口业务)","domain":"gdzjxkj.com","op":"联通","dns":"CNAME:u196510.wds168.cn","deploy":"中国联通"},{"n":"深圳市缘拓科技有限公司","assign":"刘成芳","tel":"18002560736","legal":"徐钰程","grid":"流塘置地大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道流塘社区榕树路68号流塘置地大厦1904B","ind":"金属制品、机械和设备修理","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳前海雅伦信息技术有限公司","assign":"马晓雷","tel":"18608622666","legal":"陈登","grid":"福城时代广场","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区13区宝民一路221号福城时代广场一单元305","ind":"零售业","size":"M(中型)","staff":63,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳灵犀跨境贸易有限公司","assign":"马晓雷","tel":"15218703836","legal":"张佳硕","grid":"福城时代广场","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区13区宝民一路221号福城时代广场一单元605","ind":"-","size":"S(小型)","staff":17,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)","domain":"","op":"","dns":"","deploy":""},{"n":"嘉视（深圳）电子有限公司","assign":"马晓雷","tel":"18610757372","legal":"龙文志","grid":"福城时代广场","grade":"C","src":"他网","risk":"轻涉诉","addr":"深圳市宝安区新安街道文汇社区13区宝民一路221号福城时代广场一单元606","ind":"电子零配件制造","size":"S(小型)","staff":10,"stock":"否(他网客户)","rec":"【核心】CNAME建站托管→天翼云建站迁移轻量切入；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有进出口业务)","domain":"bvtjs.com","op":"联通","dns":"CNAME:box17289798986.s802.pc51.com","deploy":"中国联通"},{"n":"深圳市伊莱斯科技有限公司","assign":"马晓雷","tel":"0755-29075495","legal":"曹一龙","grid":"鸿泰富大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区建安一路514号风尚时代D511B","ind":"科技推广和应用服务业","size":"XS(微型)","staff":9,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市弗兰格乐器有限公司","assign":"马晓雷","tel":"0755-61152201","legal":"曹一龙","grid":"鸿泰富大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区建安一路514号风尚时代D513","ind":"专业技术服务业","size":"S(小型)","staff":11,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳橙艺工贸有限公司","assign":"马晓雷","tel":"18948181242","legal":"林秋方","grid":"诺铂广场","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区前进一路269号冠利达大厦1栋926","ind":"零售业","size":"XS(微型)","staff":6,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳斯达王模具有限公司","assign":"马晓雷","tel":"13692172503","legal":"陈思健","grid":"裕安13区宝通大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道文汇社区宝民一路215号宝通大厦1405-1407","ind":"批发业","size":"XS(微型)","staff":7,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市正方行科技有限公司","assign":"潘文伟","tel":"13724247012","legal":"高玉珠","grid":"西乡街道办事处","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道办事处铁岗居委会益成工业园B栋二楼厂房","ind":"电气机械和器材制造业","size":"S(小型)","staff":25,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市昇科科技有限公司","assign":"潘文伟","tel":"13714227294","legal":"张英","grid":"燕达科技园","grade":"B","src":"他网","risk":"轻涉诉","addr":"深圳市宝安区西乡街道富华社区凤凰岗燕达工业区厂房301","ind":"电子零配件制造","size":"S(小型)","staff":39,"stock":"否(他网客户)","rec":"【核心】CNAME建站托管→天翼云建站迁移轻量切入；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有进出口业务)","domain":"sz-sk.com","op":"联通","dns":"CNAME:box17512758391.s166.pc51.com","deploy":"中国联通"},{"n":"深圳市火银顺飞供应链有限公司","assign":"钱仁伦","tel":"18908619618","legal":"史顺飞","grid":"华丰裕安商务大厦","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道海富社区45区107国道华丰裕安商务大厦409","ind":"批发业","size":"XS(微型)","staff":12,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市环家科技有限公司","assign":"钱仁伦","tel":"18822806440","legal":"张柳容","grid":"华丰裕安商务大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道海富社区45区107国道华丰裕安商务大厦522","ind":"零售业","size":"XS(微型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳锐格泰克信息技术有限公司","assign":"钱仁伦","tel":"13924636785","legal":"王同平","grid":"鸿都商务大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道海富社区45区鸿都商务大厦A栋203A","ind":"软件和信息技术服务业","size":"XS(微型)","staff":6,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市博纳德精密仪器有限公司","assign":"钱仁伦","tel":"13528837771","legal":"方军","grid":"鸿都商务大厦","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道海富社区45区鸿都商务大厦C栋505","ind":"批发业","size":"XS(微型)","staff":14,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市陇欧科技有限公司","assign":"钱仁伦","tel":"18551414096","legal":"温键龙","grid":"鸿都商务大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道海富社区45区鸿都商务大厦C栋639","ind":"零售业","size":"S(小型)","staff":9,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市抖莱咪文化传媒有限公司","assign":"钱仁伦","tel":"-","legal":"刘丽珍","grid":"七星创意工场","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道凤凰岗社区宝田一路336号七星创意工场创业楼101","ind":"文化艺术业","size":"XS(微型)","staff":9,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市地森科技有限公司","assign":"石胜","tel":"0755-36842800","legal":"程少春","grid":"好运来大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道河东社区宝安80区好运来广场河东大厦FA5-015A","ind":"软件和信息技术服务业","size":"XS(微型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市伯尼科技有限公司","assign":"石胜","tel":"18938033671","legal":"廖明栋","grid":"河东骏濠园商务楼宇部分","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道河东社区河东骏濠园1栋A单元1518","ind":"批发业","size":"XS(微型)","staff":17,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市汇德镁电子科技有限公司","assign":"石胜","tel":"13923419042","legal":"张晓峰","grid":"河东骏濠园商务楼宇部分","grade":"B","src":"他网","risk":"需关注","addr":"深圳市宝安区西乡街道河东社区河东骏濠园1栋A单元917","ind":"电子元件制造","size":"S(小型)","staff":14,"stock":"否(他网客户)","rec":"【核心】自有服务器在移动机房→迁电信IDC/天翼云+专线+固定IP一揽子；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有","domain":"huidemeipcb.com","op":"移动","dns":"A记录(直连)","deploy":"世纪互联(中国移动线路)"},{"n":"深圳市奇灵科技有限公司","assign":"石胜","tel":"18566728297","legal":"卢光慧","grid":"新城广场","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区新城广场G2196","ind":"零售业","size":"XS(微型)","staff":6,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市科扬电子科技有限公司","assign":"詹炳钦","tel":"18688997126","legal":"旷科","grid":"宝运达","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心2号厂房422","ind":"电气机械和器材制造业","size":"XS(微型)","staff":18,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市新启点智能贸易有限公司","assign":"詹炳钦","tel":"18682207313","legal":"杨宇波","grid":"宝运达","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心2号厂房423","ind":"建筑安装业","size":"XS(微型)","staff":22,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市欧泰威贸易有限公司","assign":"詹炳钦","tel":"0755-6557015","legal":"邹春晖","grid":"宝运达","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心信息大厦1609、1608","ind":"批发业","size":"S(小型)","staff":14,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳鹏胜凯科技有限公司","assign":"詹炳钦","tel":"0755-23321450","legal":"涂小连","grid":"宝运达","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心信息大厦1610","ind":"批发业","size":"S(小型)","staff":62,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市千识科技有限公司","assign":"詹炳钦","tel":"15216176013","legal":"钟金海","grid":"宝运达","grade":"B","src":"他网","risk":"干净","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心研发综合楼610","ind":"软件开发","size":"XS(微型)","staff":5,"stock":"否(他网客户)","rec":"【核心】自有服务器在移动机房→迁电信IDC/天翼云+专线+固定IP一揽子；SDWAN+云电脑；商企宽带+组网WiFi整体策反；跨境专线(有进出口业务)","domain":"szqianshi.cn","op":"移动","dns":"A记录(直连)","deploy":"世纪互联(中国移动线路)"},{"n":"深圳市滴答数字实业有限公司","assign":"詹炳钦","tel":"-","legal":"黄富","grid":"宝运达","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心综合楼二607","ind":"零售业","size":"S(小型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市宙合实业有限公司","assign":"詹炳钦","tel":"-","legal":"吴飞娥","grid":"宝运达","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区宝运达物流中心综合楼二612","ind":"零售业","size":"S(小型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市兴科达知识产权代理有限公司","assign":"詹炳钦","tel":"15999692648","legal":"徐民奎","grid":"宝运达","grade":"C","src":"他网","risk":"未扫描","addr":"深圳市宝安区西乡街道富华社区宝运达研发综合楼1103","ind":"其他科学研究和技术服务业","size":"XS(微型)","staff":5,"stock":"否(他网客户)","rec":"【核心】CNAME建站托管→天翼云建站迁移轻量切入；SDWAN+云电脑；商企宽带+组网WiFi整体策反","domain":"xkd-ip.com.cn","op":"联通","dns":"CNAME:u142382.iyz168.com","deploy":"中国联通"},{"n":"深圳市恒盛杰科技有限公司","assign":"詹炳钦","tel":"0755-27950745","legal":"张贞祥","grid":"宝运达","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区宝运达研发综合楼315","ind":"批发业","size":"XS(微型)","staff":9,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市弘翔瑞科技有限公司","assign":"詹炳钦","tel":"18688777967","legal":"张建","grid":"宝运达","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道富华社区宝运达研发综合楼903","ind":"零售业","size":"XS(微型)","staff":6,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市芯联芯电子有限公司","assign":"詹炳钦","tel":"13924671121","legal":"肖志强","grid":"桃花源科技创新园","grade":"C","src":"他网","risk":"干净","addr":"深圳市宝安区西乡街道桃花源科技创新园蚝业分园A栋厂房4B","ind":"光学光电子制造","size":"S(小型)","staff":14,"stock":"否(他网客户)","rec":"【核心】CNAME建站托管→天翼云建站迁移轻量切入；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有进出口业务)","domain":"china-modules.com","op":"联通","dns":"CNAME:u262683.wds168.cn","deploy":"中国联通"},{"n":"深圳市乐技网络科技有限公司","assign":"詹炳钦","tel":"13428768109","legal":"龙思萍","grid":"雍华源A栋商务大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道宝田二路雍华源C栋1单元1104","ind":"零售业","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市中云智慧科技有限公司","assign":"詹炳钦","tel":"13751114206","legal":"潘希敏","grid":"雍华源A栋商务大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区宝田二路6号雍华源A栋商务楼2110-2111","ind":"批发业","size":"XS(微型)","staff":9,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市鹏诚新能源科技有限公司","assign":"詹炳钦","tel":"13684924995","legal":"谭祖宪","grid":"雍华源A栋商务大厦","grade":"A","src":"他网","risk":"轻涉诉","addr":"深圳市宝安区西乡街道臣田社区宝田二路6号雍华源A栋商务楼701-702","ind":"电池制造","size":"S(小型)","staff":21,"stock":"否(他网客户)","rec":"【核心】自有服务器在移动机房→迁电信IDC/天翼云+专线+固定IP一揽子；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有","domain":"pchne.com","op":"移动","dns":"A记录(直连)","deploy":"世纪互联(中国移动线路); 中国电信"},{"n":"广东腾跃通电缆有限公司","assign":"詹炳钦","tel":"15019248080","legal":"吴振保","grid":"雍华源A栋商务大厦","grade":"B","src":"他网","risk":"轻涉诉","addr":"深圳市宝安区西乡街道臣田社区宝田二路6号雍华源B栋918","ind":"电线、电缆及其他电工器材","size":"XS(微型)","staff":1,"stock":"否(他网客户)","rec":"【核心】自有服务器在联通机房→迁电信IDC/天翼云+专线+固定IP一揽子；专线+固定IP(ERP/外贸/设备运维)；商企宽带+组网WiFi整体策反；跨境专线(有","domain":"tytcable.com","op":"联通","dns":"A记录(直连)","deploy":"中国联通"},{"n":"深圳市霍尼斯特科技有限公司","assign":"张集峰","tel":"0755-61152718","legal":"康家明","grid":"富源工业园","grade":"A","src":"跨境","risk":"轻涉诉","addr":"深圳市宝安区新安街道海富社区45区翻身路富源工业区1栋富源大厦302A","ind":"科技推广和应用服务业","size":"S(小型)","staff":10,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市四海行科技有限公司","assign":"张集峰","tel":"15986698647","legal":"龚胜群","grid":"富源工业园","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区新安街道海富社区45区翻身路富源工业区1栋富源大厦313","ind":"批发业","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市睿佳利电子商务有限公司","assign":"张集峰","tel":"13530832118","legal":"杨天育","grid":"华创达商务大厦","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道46区翠竹路海富社区华创达商务大厦201","ind":"批发业","size":"XS(微型)","staff":10,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市特佳安科技有限公司","assign":"张集峰","tel":"0755-29446392","legal":"刘磊","grid":"华创达怡景大厦","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区45区107国道旁华创达怡景大厦701室（办公场所）","ind":"计算机、通信和其他电子设","size":"XS(微型)","staff":5,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市粤浩研磨材料有限公司","assign":"张集峰","tel":"18576611087","legal":"胡志","grid":"华丰新安商务大厦","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区新安街道海富社区45区怡景大厦A栋华丰新安商务大厦广深路新安段259（109）","ind":"批发业","size":"XS(微型)","staff":12,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市美盛泰科技有限公司","assign":"钟宏胜","tel":"18938027234","legal":"裴雄燃","grid":"臣田工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区宝民二路190号臣田工业区商业楼8533","ind":"科技推广和应用服务业","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳沃圣数码科技有限公司","assign":"钟宏胜","tel":"13480128049","legal":"刘爱添","grid":"臣田工业区","grade":"C","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区宝田二路2号臣田工业区28栋4层","ind":"计算机、通信和其他电子设","size":"S(小型)","staff":27,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市易之酷科技有限公司","assign":"钟宏胜","tel":"18123673715","legal":"李冉","grid":"臣田工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区臣田定军山电影文化科技创意园B栋701","ind":"批发业","size":"XS(微型)","staff":7,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市恒鸿高科技术有限公司","assign":"钟宏胜","tel":"13823302301","legal":"刘刚","grid":"臣田工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区臣田工业区25栋201","ind":"批发业","size":"XS(微型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市恰友科技有限公司","assign":"钟宏胜","tel":"0755-23595295","legal":"许传文","grid":"臣田工业区","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道臣田社区臣田工业区25栋4层","ind":"软件和信息技术服务业","size":"S(小型)","staff":24,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市诚意枫光电有限公司","assign":"钟宏胜","tel":"18923464931","legal":"周丹苹","grid":"臣田工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区臣田工业区30栋3层","ind":"计算机、通信和其他电子设","size":"S(小型)","staff":61,"stock":"是","rec":"跨境专线(店铺/平台/ERP加速)、专线+固定IP(ERP/订单)","domain":"","op":"","dns":"","deploy":""},{"n":"深圳市怡嘉淇科技有限公司","assign":"钟宏胜","tel":"13632963986","legal":"陶水仙","grid":"臣田工业区","grade":"A","src":"跨境","risk":"干净","addr":"深圳市宝安区西乡街道臣田社区臣田工业区32栋4层","ind":"批发业","size":"S(小型)","staff":19,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、商企宽带+固定IP","domain":"","op":"","dns":"","deploy":""},{"n":"深芯链科技（深圳）有限公司","assign":"钟宏胜","tel":"13262650296","legal":"秦秀琼","grid":"臣田工业区","grade":"B","src":"跨境","risk":"未深查","addr":"深圳市宝安区西乡街道臣田社区臣田工业区36栋401","ind":"软件和信息技术服务业","size":"S(小型)","staff":7,"stock":"否","rec":"跨境专线(店铺/平台/ERP加速)、专线+云电脑(SDWAN试点)","domain":"","op":"","dns":"","deploy":""}],"users":["詹炳钦","刘成芳","江静竹","钟宏胜","马晓雷","钱仁伦","非裕安","黄锋","张集峰","石胜","潘文伟"],"tokUsers":{"9y4g5lon":"詹炳钦","xon664k8":"刘成芳","rgvl33bh":"江静竹","pd9h2mjq":"钟宏胜","3hiorzfy":"马晓雷","f4d8k3gp":"非裕安","8ohax2e9":"黄锋","uythmvy6":"钱仁伦","kl1umm6x":"张集峰","50t5dsm9":"石胜","1akkj1ms":"潘文伟"},"tokAdmin":"admf01jp58156"}""")
RECS = DATA['recs']
USERS = DATA['users']
TOK_USERS = DATA['tokUsers']
TOK_ADMIN = DATA['tokAdmin']

VISIT_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>裕安 · 走访登记</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Microsoft YaHei',sans-serif;background:#eef2f7;color:#333;line-height:1.55;font-size:14px}
.wrap{max-width:1000px;margin:0 auto;padding:10px}
.hd{background:linear-gradient(135deg,#0a4da3,#1c86e8);color:#fff;border-radius:12px;padding:16px 20px 13px}
.hd h1{font-size:19px}
.hd .sub{font-size:12px;opacity:.92;margin-top:4px}
.kpiRow{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:10px 0}
.kpi{background:#fff;border-radius:10px;padding:9px 4px;text-align:center;box-shadow:0 1px 4px rgba(10,77,163,.08);border-left:4px solid #0a4da3}
.kv{font-size:20px;font-weight:800;color:#0a4da3}
.kv.grn{color:#16a085}.kv.red{color:#c0392b}.kv.org{color:#e67e22}
.kl{font-size:11px;color:#555}
.card{background:#fff;border-radius:12px;padding:12px 14px;margin-bottom:10px;box-shadow:0 1px 5px rgba(10,77,163,.07)}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{border:1.5px solid #b9d4ee;border-radius:15px;padding:3px 11px;font-size:12px;color:#0a4da3;cursor:pointer;background:#fff;user-select:none}
.chip.on{background:#0a4da3;color:#fff;border-color:#0a4da3;font-weight:700}
.srow input{width:100%;border:1.5px solid #b9d4ee;border-radius:8px;padding:7px 11px;font-size:13px;outline:none;margin-top:8px}
table{width:100%;border-collapse:collapse;font-size:12.2px}
th{background:#0a4da3;color:#fff;padding:6px 4px;text-align:center;white-space:nowrap}
td{padding:6px 4px;border-bottom:1px solid #e8eef5;text-align:center}
tbody tr{cursor:pointer}
tbody tr:hover td{background:#f0f7ff}
.tName{text-align:left;font-weight:600;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.badge{display:inline-block;border-radius:10px;padding:1px 8px;font-size:10.5px;font-weight:700}
.b-org{background:#fdf3e7;color:#b96a00}.b-red{background:#fdecea;color:#c0392b}.b-grn{background:#e9f7ef;color:#1e8449}.b-yel{background:#fef9e7;color:#b7950b}.b-blu{background:#eaf3fb;color:#0a4da3}.b-gry{background:#f0f0f0;color:#777}
.btnReg{background:#0a4da3;color:#fff;border:none;border-radius:6px;padding:4px 10px;font-size:11.5px;cursor:pointer}
.btnXl{background:#16a085;color:#fff;border:none;border-radius:8px;padding:9px 18px;font-size:13.5px;cursor:pointer;font-weight:700}
.hint{font-size:11.5px;color:#999;margin:6px 0}
.pGrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px;margin-bottom:10px}
.pCard{border:1px solid #dbe6f2;border-radius:9px;padding:8px 10px;background:#fff}
.pCard .pn{font-weight:700;color:#0a4da3;font-size:13.5px}
.pCard .pd{font-size:11px;color:#888;margin-top:2px}
.bar{height:6px;background:#e8eef5;border-radius:3px;margin-top:5px;overflow:hidden}
.bar i{display:block;height:100%;background:#16a085;border-radius:3px}
.overlay{position:fixed;inset:0;background:rgba(10,30,60,.55);display:none;z-index:100;align-items:flex-start;justify-content:center;padding:14px 8px;overflow:auto}
.overlay.on{display:flex}
.modal{background:#fff;border-radius:12px;max-width:560px;width:100%;padding:14px 16px;box-shadow:0 8px 40px rgba(0,0,0,.3)}
.mclose{float:right;font-size:19px;color:#999;cursor:pointer}
.modal h3{color:#0a4da3;font-size:15px;margin-bottom:3px;padding-right:22px}
.mtags{margin:3px 0 8px}
.minfo{font-size:12.3px;color:#444;border:1px solid #e0eaf5;border-radius:8px;padding:8px 10px;margin-bottom:9px;background:#f8fbfe}
.form label{display:block;font-size:12px;color:#0a4da3;font-weight:700;margin:9px 0 4px}
.form input[type=date],.form input[type=text],.form textarea{width:100%;border:1.5px solid #cfe0f0;border-radius:8px;padding:7px 10px;font-size:13.5px;outline:none;font-family:inherit}
.form textarea{height:56px;resize:vertical}
.seg{display:flex;flex-wrap:wrap;gap:6px}
.seg label{margin:0;cursor:pointer;border:1.5px solid #cfe0f0;border-radius:15px;padding:5px 12px;font-size:12.5px;color:#333;font-weight:400;background:#fff}
.seg input{display:none}
.seg label:has(input:checked){background:#0a4da3;border-color:#0a4da3}
.seg label:has(input:checked) span{color:#fff}
.pk label:has(input:checked){background:#16a085;border-color:#16a085}
.form .btnGo{width:100%;margin-top:13px;background:#0a4da3;color:#fff;border:none;border-radius:9px;padding:11px;font-size:15px;font-weight:700;cursor:pointer}
.toast{position:fixed;left:50%;bottom:34px;transform:translateX(-50%);background:#16a085;color:#fff;padding:9px 22px;border-radius:20px;font-size:13px;display:none;z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.25)}
.ft{text-align:center;color:#8898aa;font-size:10.5px;padding:10px 0 18px}
.gtCard{background:#fff;border-radius:12px;padding:14px 16px;margin-bottom:12px;box-shadow:0 1px 5px rgba(10,77,163,.07);display:none}
.gtCard.on{display:block}
.gtHead{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px}
.gtTitle{font-size:15px;font-weight:700;color:#0a4da3}
.gtWrap{overflow-x:auto}
.gantt{width:100%;border-collapse:collapse;font-size:11.5px;min-width:640px}
.gantt th{background:#0a4da3;color:#fff;padding:5px 4px;text-align:center;position:sticky;top:0;white-space:nowrap}
.gantt td{padding:4px;border-bottom:1px solid #e8eef5;white-space:nowrap}
.gUser{font-weight:700;color:#333;cursor:pointer}
.gUser:hover{color:#0a4da3}
.gTrack{position:relative;height:20px;background:#eef3fa;border-radius:10px;overflow:hidden;width:170px}
.gBarDone{position:absolute;left:0;top:0;bottom:0;background:linear-gradient(90deg,#16a085,#2ec98f);border-radius:10px;display:flex;align-items:center;justify-content:flex-end;padding-right:6px;color:#fff;font-size:10px;font-weight:700}
.gBarPlan{position:absolute;left:0;top:0;bottom:0;border-left:2px dashed #b7950b}
.gPct{font-weight:700;color:#16a085}
.gA{color:#c0392b;font-weight:700}
.gB{color:#e67e22;font-weight:700}
.gC{color:#b7950b}.gD{color:#999}.gE{color:#aaa}
.gRank{font-weight:800;color:#0a4da3}
.opCard{background:#fff;border-radius:12px;padding:14px 16px;margin-bottom:12px;box-shadow:0 1px 5px rgba(10,77,163,.07);display:none}
.opCard.on{display:block}
.opTabs{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}
.opTab{padding:5px 14px;border-radius:16px;font-size:12.5px;cursor:pointer;background:#e3edf8;color:#0a4da3;font-weight:600}
.opTab.on{background:#0a4da3;color:#fff}
.opTbl{width:100%;border-collapse:collapse;font-size:12.3px}
.opTbl th{background:#0a4da3;color:#fff;padding:6px 5px}
.opTbl td{padding:5px;border-bottom:1px solid #e8eef5}
.opTbl tbody tr{cursor:pointer}
.opEmpty{color:#999;text-align:center;padding:14px;font-size:12.5px}
.hist{font-size:11.5px;color:#888;background:#f5f9fd;border-radius:7px;padding:6px 9px;margin-bottom:8px}
#gate{display:none;text-align:center;padding:60px 20px}
#gate .ic{font-size:44px}
#gate h2{color:#0a4da3;margin:10px 0 6px}
#gate p{color:#888;font-size:13px}
@media(max-width:640px){.kpiRow{grid-template-columns:repeat(4,1fr)}.hd h1{font-size:17px}.tName{max-width:110px}}
</style></head><body>
<div class="wrap">
<div id="gate"><div class="ic">🔒</div><h2>链接无效或已过期</h2><p>请联系中心管理员（吴局）获取您的专属走访链接</p></div>

<div id="app" style="display:none">
<div class="hd">
<h1 id="hdTitle">裕安 · 走访登记</h1>
<div class="sub" id="hdSub"></div>
</div>
<div class="kpiRow" id="kpiRow"></div>

<div class="card" id="adminTop" style="display:none">
<button class="btnXl" onclick="window.open(EXPURL)">📥 导出Excel（含回填）</button>
<button class="btnXl" style="background:#e67e22" onclick="genReport()">📣 生成战报</button>
<div class="hint">导出=84家+回填列（含到期/跟进） · 战报=一键生成群发文本（点弹窗复制）</div>
</div>

<div class="card" style="padding:8px 14px"><a href="./sales.html" target="_blank" style="color:#0a4da3;font-weight:700;font-size:13px">📖 话术速查手册（上门/电话话术，点开即看）</a></div>
<div class="card">
<div class="chips" id="chips"></div>
<div class="srow"><input id="q" placeholder="搜索企业名/电话/网格（回车）"></div>
</div>

<div class="card" id="adminUsers" style="display:none">
<div class="hint">各分配人回填进度（点击查看该人客户）：</div>
<div class="pGrid" id="pGrid"></div>
</div>

<div class="gtCard" id="ganttCard">
<div class="gtHead">
<div class="gtTitle">📊 走访进度甘特图</div>
<div style="display:flex;gap:6px;align-items:center">
<span style="font-size:11px;color:#888">目标：A级15日前 / 全量80%月底</span>
<button class="btnReg" style="padding:5px 12px;font-size:12px" onclick="loadResults('ganttRefreshed')">🔄 刷新进度</button>
</div>
</div>
<div class="gtWrap"><table class="gantt" id="gtTbl"></table></div>
<div class="hint" id="gtHint"></div>
</div>

<div class="opCard" id="oppCard">
<div class="gtHead">
<div class="gtTitle">🎯 商机汇总（点击行查看客户）</div>
<button class="btnReg" style="padding:5px 12px;font-size:12px" onclick="loadResults('oppRefreshed')">🔄 刷新</button>
</div>
<div class="opTabs">
<div class="opTab on" onclick="setOpp('all',this)">全部</div>
<div class="opTab" onclick="setOpp('A',this)">✅ 已成交</div>
<div class="opTab" onclick="setOpp('B',this)">🔥 高意向</div>
</div>
<div style="overflow:auto"><table class="opTbl" id="oppTbl"></table></div>
</div>

<div class="card">
<div class="hint" id="tbHint"></div>
<div style="overflow:auto"><table id="tbl"></table></div>
</div>

<div class="ft">裕安政商 · 走访登记系统 · 提交即回填中心台账 · 局长端实时可见</div>
</div>
</div>

<div class="overlay" id="ov" onclick="if(event.target===this)closeM()">
<div class="modal">
<span class="mclose" onclick="closeM()">✕</span>
<div id="mBody"></div>
</div>
</div>
<div class="toast" id="toast"></div>

<script>
var QS = new URLSearchParams(location.search);
var MYTOK = QS.get('u') || QS.get('k') || '';
var IS_ADMIN = false;
var R = [], USERS = [], USER = '', RS = [], latest = {};
var F = {grade:[], fill:'', risk:'', src:''};
var Q = '';
var curFilterUser = '';
var GENDATE = '';
var DATAURL = '/api/data?';
var EXPURL = '#';

function fmt(v){return String(v).replace(/\\B(?=(\\d{3})+(?!\\d))/g,',');}
function toast(m){var t=document.getElementById('toast');t.textContent=m;t.style.display='block';setTimeout(function(){t.style.display='none';},1800);}
function showGate(){ document.getElementById('gate').style.display='block'; document.getElementById('app').style.display='none'; }

function myRecs(){
  return R.filter(function(r){
    if(IS_ADMIN){ if(curFilterUser && r.assign!==curFilterUser) return false; }
    if(F.grade.length && F.grade.indexOf(r.grade)<0) return false;
    if(F.src && r.src!==F.src) return false;
    if(F.risk && r.risk!==F.risk) return false;
    if(F.fill==='todo' && latest[r.n]) return false;
    if(F.fill==='done' && !latest[r.n]) return false;
    if(F.exp30){
      var l=latest[r.n]; var today=new Date(); var d30=new Date(); d30.setDate(today.getDate()+30);
      if(!l || !l.expire) return false;
      var ex=new Date(l.expire.replace('-','/'));
      if(!(ex>=today && ex<=d30)) return false;
    }
    if(Q){ var hay=(r.n+' '+r.tel+' '+r.grid+' '+r.legal).toLowerCase(); if(hay.indexOf(Q.toLowerCase())<0) return false; }
    return true;
  });
}

function fillBadge(r){
  var l = latest[r.n];
  if(!l) return '<span class="badge b-gry">待走访</span>';
  var map = {'A':'b-grn','B':'b-blu','C':'b-yel','D':'b-gry','E':'b-gry'};
  var txt = {'A':'已成交','B':'高意向','C':'观望','D':'无意向','E':'未接触'}[l.result] || l.result;
  var t = new Date(); var tstr = t.getFullYear()+'-'+String(t.getMonth()+1).padStart(2,'0')+'-'+String(t.getDate()).padStart(2,'0');
  var b = '<span class="badge '+(map[l.result]||'b-gry')+'">'+txt+'</span>';
  if(l.expire){ b += '<br><span style="font-size:10px;color:'+(l.expire<=tstr?'#c0392b':'#888')+'">到期'+l.expire+'</span>'; }
  if(l.next && l.next < tstr){ b += ' <span class="badge b-red">跟进逾期</span>'; }
  return b;
}

function renderHead(){
  var done = R.filter(function(r){return latest[r.n];}).length;
  var a = R.filter(function(r){return r.grade==='A';}).length;
  document.getElementById('hdTitle').textContent = IS_ADMIN ? '裕安 · 走访回填总览（管理层）' : USER + ' · 走访客户包';
  document.getElementById('hdSub').textContent = IS_ADMIN ? '84家客户 · 实时回填进度 · '+GENDATE : '共'+R.length+'家客户（跨境'+R.filter(function(r){return r.src==='跨境';}).length+'家 · 他网策反'+R.filter(function(r){return r.src==='他网';}).length+'家）· 已回填'+done+'家 · 点击客户行登记走访结果';
  document.getElementById('kpiRow').innerHTML =
    '<div class="kpi"><div class="kv">'+R.length+'</div><div class="kl">'+(IS_ADMIN?'总客户':'我的客户')+'</div></div>'
    +'<div class="kpi"><div class="kv org">'+a+'</div><div class="kl">A级重点</div></div>'
    +'<div class="kpi"><div class="kv grn">'+done+'</div><div class="kl">已回填</div></div>'
    +'<div class="kpi"><div class="kv red">'+(R.length-done)+'</div><div class="kl">待走访</div></div>';
  document.getElementById('adminTop').style.display = IS_ADMIN?'':'none';
  document.getElementById('adminUsers').style.display = IS_ADMIN?'':'none';
  document.getElementById('ganttCard').classList.toggle('on', IS_ADMIN);
  document.getElementById('oppCard').classList.toggle('on', IS_ADMIN);
}

function renderUsers(){
  if(!IS_ADMIN) return;
  var s='';
  USERS.forEach(function(u){
    var mine = R.filter(function(r){return r.assign===u;});
    var done = mine.filter(function(r){return latest[r.n];}).length;
    var pct = mine.length?Math.round(done/mine.length*100):0;
    s+='<div class="pCard" style="cursor:pointer" onclick="byUser(\\''+u+'\\')"><div class="pn">'+u+'</div><div class="pd">'+mine.length+'家 · 已填'+done+' · A'+mine.filter(function(r){return r.grade==='A';}).length+'</div><div class="bar"><i style="width:'+pct+'%"></i></div></div>';
  });
  document.getElementById('pGrid').innerHTML=s;
}
function byUser(u){ curFilterUser=(curFilterUser===u?'':u); renderUsers(); render(); toast(curFilterUser?('已筛选：'+u):'已清人筛'); }

function render(){
  var list = myRecs();
  var s='<thead><tr>'+(IS_ADMIN?'<th>分配人</th>':'')+'<th>企业名称</th><th>类型</th><th>级</th><th>网格</th><th>电话</th><th>风险</th><th>走访状态</th><th>操作</th></tr></thead><tbody>';
  list.forEach(function(r,i){
    r._i=i;
    var gd=r.grade==='A'?'<span class="badge b-red">A</span>':(r.grade==='B'?'<span class="badge b-yel">B</span>':'<span class="badge b-gry">C</span>');
    var rk=r.risk==='干净'?'<span class="badge b-grn">净</span>':(r.risk==='需关注'?'<span class="badge b-red">关</span>':(r.risk==='轻涉诉'?'<span class="badge b-yel">诉</span>':'<span class="badge b-gry">—</span>'));
    s+='<tr onclick="drill('+i+')">'+(IS_ADMIN?'<td style="font-size:11px">'+r.assign+'</td>':'')+'<td class="tName">'+r.n+'</td><td><span class="badge '+(r.src==='他网'?'b-org':'b-blu')+'">'+r.src+'</span></td><td>'+gd+'</td><td style="font-size:11px">'+r.grid+'</td><td style="font-size:11px">'+r.tel+'</td><td>'+rk+'</td><td>'+fillBadge(r)+'</td><td><button class="btnReg" onclick="event.stopPropagation();drill('+i+')">登记</button></td></tr>';
  });
  s+='</tbody>';
  document.getElementById('tbl').innerHTML=s;
  document.getElementById('tbHint').innerHTML='当前 <b>'+list.length+'</b> 家'+(IS_ADMIN?'（点击行查看回填详情）':'（点击行登记走访结果）');
}

function chips(){
  var s='';
  [['grade','优先级',['A','B','C']],['src','来源',['跨境','他网']],['risk','风险',['干净','轻涉诉','需关注']]].forEach(function(g){
    s+='<span style="font-size:11px;color:#0a4da3;font-weight:700;margin-right:3px">'+g[1]+'</span>';
    g[2].forEach(function(v){ s+='<span class="chip'+(F[g[0]]===v?' on':'')+'" onclick="tgl(\\''+g[0]+'\\',\\''+v+'\\')">'+v+'</span>'; });
  });
  s+='<span style="font-size:11px;color:#0a4da3;font-weight:700;margin:0 3px 0 8px">回填</span>';
  s+='<span class="chip'+(F.fill==='done'?' on':'')+'" onclick="tglFill(\\'done\\')">已回填</span>';
  s+='<span class="chip'+(F.fill==='todo'?' on':'')+'" onclick="tglFill(\\'todo\\')">待走访</span>';
  s+='<span class="chip'+(F.exp30?' on':'')+'" onclick="F.exp30=!F.exp30;render()">30天内到期</span>';
  if(IS_ADMIN) s+='<span class="chip'+(curFilterUser?' on':'')+'" onclick="curFilterUser=\\'\\';renderUsers();render();">清人筛</span>';
  document.getElementById('chips').innerHTML=s;
}
function tgl(k,v){ F[k]=(F[k]===v?'':v); render(); }
function tglFill(v){ F.fill=(F.fill===v?'':v); render(); }

var curIdx = -1;
function drill(i){
  curIdx = i;
  var r = myRecs()[i];
  var l = latest[r.n];
  var hist = RS.filter(function(x){return x.name===r.n;});
  var histHtml = '';
  if(hist.length){
    var h = hist.slice(-3).reverse();
    histHtml = '<div class="hist">最近登记：'+h.map(function(x){return (x.date)+' '+(x.user)+' ['+(x.result)+'] '+(x.products||[]).join('/').slice(0,20);}).join('<br>')+'</div>';
  }
  var b='<h3>'+r.n+'</h3>'
   +'<div class="mtags"><span class="badge '+(r.src==='他网'?'b-yel':'b-blu')+'">'+r.src+'</span> <span class="badge '+(r.grade==='A'?'b-red':'b-yel')+'">'+r.grade+'级</span> <span class="badge b-gry">'+r.grid+'</span> <span class="badge b-gry">'+r.size+'·'+r.staff+'人</span></div>'
   +'<div class="minfo">联系人：<b>'+r.legal+'</b> <a href="tel:'+r.tel+'" style="color:#0a4da3;font-weight:700">📞 '+r.tel+'</a><br>地址：'+r.addr+(r.src==='他网'&&r.domain?'<br><b style="color:#b96a00">官网弹药</b>：'+r.domain+'（部署在<b>'+r.op+'</b>'+(r.dns?' · '+r.dns:'')+(r.deploy?' · '+r.deploy:'')+'）→ 按手册“服务器搬家不是重买”打法谈':'')+'<br>建议：'+(r.rec||'-')+'</div>'
   + histHtml;
  if(IS_ADMIN){ document.getElementById('mBody').innerHTML=b; document.getElementById('ov').classList.add('on'); return; }
  var today = new Date(); var td = today.getFullYear()+'-'+String(today.getMonth()+1).padStart(2,'0')+'-'+String(today.getDate()).padStart(2,'0');
  b+='<div class="form">'
   +'<label>📅 走访日期</label><input type="date" id="fDate" value="'+td+'">'
   +'<label>📌 走访结果</label><div class="seg" id="fSeg">'
   +'<label><input type="radio" name="rst" value="A" '+(l&&l.result==='A'?'checked':'')+'><span>A 已成交</span></label>'
   +'<label><input type="radio" name="rst" value="B" '+(l&&l.result==='B'?'checked':'')+'><span>B 高意向</span></label>'
   +'<label><input type="radio" name="rst" value="C" '+(l&&l.result==='C'?'checked':'')+'><span>C 观望</span></label>'
   +'<label><input type="radio" name="rst" value="D" '+(l&&l.result==='D'?'checked':'')+'><span>D 无意向</span></label>'
   +'<label><input type="radio" name="rst" value="E" '+(l&&l.result==='E'?'checked':'')+'><span>E 未接触上</span></label>'
   +'</div>'
   +'<label>🎯 意向产品（可多选）</label><div class="seg pk" id="fPk">'
   +['跨境专线','SD-WAN','固定IP','商企宽带','组网WiFi','火情AI','他网迁移','天翼云'].map(function(p){
      var chk = (l && (l.products||[]).indexOf(p)>=0)?'checked':'';
      return '<label><input type="checkbox" name="pk" value="'+p+'" '+chk+'><span>'+p+'</span></label>';
   }).join('')
   +'</div>'
   +'<label>📅 客户宽带到期时间（现场问到·策反黄金锚点）</label><input type="date" id="fExpire" value="'+(l&&l.expire?l.expire:'')+'">'
   +'<label>📝 备注（现场情况/客户关注点）</label><textarea id="fNote">'+(l?l.note:'')+'</textarea>'
   +'<label>⏰ 下次跟进时间（日期）</label><input type="date" id="fNext" value="'+(l&&l.next?l.next:'')+'">'
   +'<button class="btnGo" onclick="submit()">✅ 提交走访结果</button>'
   +'</div>';
  document.getElementById('mBody').innerHTML=b;
  document.getElementById('ov').classList.add('on');
}

function submit(){
  var r = myRecs()[curIdx];
  var rst = document.querySelector('input[name=rst]:checked');
  if(!rst){ toast('请选择走访结果'); return; }
  var pks = Array.from(document.querySelectorAll('input[name=pk]:checked')).map(function(x){return x.value;});
  var body = {
    tok: MYTOK,
    name: r.n,
    date: document.getElementById('fDate').value,
    result: rst.value,
    products: pks,
    note: document.getElementById('fNote').value.trim(),
    expire: document.getElementById('fExpire').value.trim(),
    next: document.getElementById('fNext').value.trim()
  };
  fetch('/api/submit', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)})
   .then(function(x){return x.json();})
   .then(function(j){
     if(j.ok){ closeM(); toast('✅ 已提交，中心台账已更新'); loadResults(); }
     else toast('提交失败：'+(j.err||''));
   })
   .catch(function(e){ toast('网络异常，请重试'); });
}

var OPP_FILTER = 'all';
function setOpp(v, el){
  OPP_FILTER = v;
  document.querySelectorAll('.opTab').forEach(function(t){t.classList.remove('on');});
  if(el) el.classList.add('on');
  renderOpp();
}
function renderGantt(){
  if(!IS_ADMIN) return;
  var rows = [];
  USERS.forEach(function(u){
    var mine = R.filter(function(r){return r.assign===u;});
    var done = mine.filter(function(r){return latest[r.n];}).length;
    var pct = mine.length ? Math.round(done/mine.length*100) : 0;
    var cnt = {A:0,B:0,C:0,D:0,E:0};
    mine.forEach(function(r){
      var l = latest[r.n]; if(l && cnt[l.result]!==undefined) cnt[l.result]++;
    });
    var aPlan = mine.filter(function(r){return r.grade==='A';});
    var aDone = aPlan.filter(function(r){return latest[r.n];}).length;
    var aPct = aPlan.length ? Math.round(aDone/aPlan.length*100) : 0;
    var prods = {};
    mine.forEach(function(r){
      var l = latest[r.n]; if(l && l.products) l.products.forEach(function(p){ prods[p]=(prods[p]||0)+1; });
    });
    var top = Object.keys(prods).sort(function(a,b){return prods[b]-prods[a];}).slice(0,2).map(function(p){return p+'x'+prods[p];}).join(' ');
    rows.push({u:u, total:mine.length, done:done, pct:pct, cnt:cnt, aPlan:aPlan.length, aDone:aDone, aPct:aPct, top:top});
  });
  rows.sort(function(a,b){ return (b.pct-a.pct)||(b.cnt.A-a.cnt.A)||(b.cnt.B-a.cnt.B); });
  var totalDone = 0, totalAll = 0, totalA=0, totalB=0;
  rows.forEach(function(x){ totalDone+=x.done; totalAll+=x.total; totalA+=x.cnt.A; totalB+=x.cnt.B; });
  var d15 = new Date('2026-11-15'); var now = new Date();
  var s = '<thead><tr><th>排名</th><th>分配人</th><th>总进度(已回填/客户)</th><th>进度条</th><th>A级阶段<br>(15日前100%)</th><th style="color:#ffd54f">成交</th><th style="color:#ffb74d">高意向</th><th>观望</th><th>意向产品TOP</th></tr></thead><tbody>';
  var medal = ['🥇','🥈','🥉'];
  rows.forEach(function(x,i){
    var barW = Math.min(x.pct,100);
    var aState = x.aPct>=100 ? 'background:#1e8449;color:#fff' : (x.aPct>=50 ? 'background:#fef9e7;color:#b7950b' : 'background:#fdecea;color:#c0392b');
    s+='<tr>'
     +'<td class="gRank">'+(i<3?medal[i]:'')+((i<3?'':'')+(i+1))+'</td>'
     +'<td class="gUser" onclick="byUser(\\''+x.u+'\\')">'+x.u+'</td>'
     +'<td><b>'+x.done+'</b>/'+x.total+'</td>'
     +'<td><div class="gTrack"><div class="gBarDone" style="width:'+barW+'%">'+(barW>=14?barW+'%':'')+'</div></div> <span class="gPct">'+barW+'%</span></td>'
     +'<td style="'+aState+';border-radius:9px">'+x.aDone+'/'+x.aPlan+'·'+x.aPct+'%</td>'
     +'<td class="gA">'+(x.cnt.A||'-')+'</td>'
     +'<td class="gB">'+(x.cnt.B||'-')+'</td>'
     +'<td class="gC">'+(x.cnt.C||'-')+'</td>'
     +'<td style="font-size:10.5px;color:#555;white-space:normal;max-width:130px">'+(x.top||'-')+'</td>'
     +'</tr>';
  });
  s+='</tbody>';
  document.getElementById('gtTbl').innerHTML = s;
  document.getElementById('gtHint').innerHTML = '全员：回填 <b>'+totalDone+'/'+totalAll+'</b>（'+(totalAll?Math.round(totalDone/totalAll*100):0)+'%）· 成交<b style="color:#c0392b">'+totalA+'</b> · 高意向<b style="color:#e67e22">'+totalB+'</b> · '+($(d15>now?'':'color:#999'))+'节点：11.15前A级100%、11.30前全量80%。点击人名=筛选该人客户';
}
function renderOpp(){
  if(!IS_ADMIN) return;
  var list = RS.filter(function(x){
    if(OPP_FILTER==='all') return x.result==='A'||x.result==='B';
    return x.result===OPP_FILTER;
  });
  var recMap = {};
  R.forEach(function(r){ recMap[r.n]=r; });
  var latest4name = {};
  RS.forEach(function(x){ latest4name[x.name]=x; });
  var seen = {}; var dedup = [];
  for(var i=list.length-1;i>=0;i--){ var x=list[i]; if(!seen[x.name]){ seen[x.name]=1; dedup.push(x);} }
  dedup.reverse();
  var RT={A:'✅已成交',B:'🔥高意向'};
  var s='<thead><tr><th>结果</th><th>企业</th><th>分配人</th><th>意向产品</th><th>客户到期</th><th>下次跟进</th><th>备注</th></tr></thead><tbody>';
  if(!dedup.length){ s+='</tbody>'; document.getElementById('oppTbl').innerHTML=s+'<div class="opEmpty">暂无'+(OPP_FILTER==='all'?'成交/高意向':RT[OPP_FILTER])+'记录</div>'; return; }
  dedup.forEach(function(x){
    var r = recMap[x.name]||{};
    var name = x.name;
    var l = latest4name[x.name];
    s+='<tr onclick="drillByName('+JSON.stringify(name)+')">'
     +'<td><span class="badge '+(x.result==='A'?'b-grn':'b-yel')+'">'+(RT[x.result]||x.result)+'</span></td>'
     +'<td style="text-align:left;font-weight:600;max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+name+'</td>'
     +'<td>'+x.user+'</td>'
     +'<td style="font-size:11px">'+(l&&l.products?l.products.join('、'):'-')+'</td>'
     +'<td style="font-size:11px;color:'+(l&&l.expire&&l.expire<=todayStr()?'#c0392b':'#333')+'">'+(l&&l.expire?l.expire:'-')+'</td>'
     +'<td style="font-size:11px;color:'+(l&&l.next&&l.next<todayStr()?'#c0392b':'#333')+'">'+(l&&l.next?l.next:'-')+'</td>'
     +'<td style="font-size:11px;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left">'+(x.note||'-')+'</td>'
     +'</tr>';
  });
  s+='</tbody>';
  document.getElementById('oppTbl').innerHTML = s;
}
function todayStr(){ var d=new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
function drillByName(n){ curFilterUser=''; var i=R.findIndex(function(r){return r.n===n;}); if(i>=0){ drill(i); } }
function genReport(){
  var done = R.filter(function(r){return latest[r.n];}).length;
  var total = R.length;
  var cnt = {A:0,B:0,C:0,D:0,E:0};
  var perUser = {}, prod = {};
  RS.forEach(function(x){
    if(cnt[x.result]!==undefined) cnt[x.result]++;
    perUser[x.user] = (perUser[x.user]||0)+1;
    (x.products||[]).forEach(function(p){ prod[p]=(prod[p]||0)+1; });
  });
  var d7 = new Date(); d7.setDate(d7.getDate()-7);
  var week = RS.filter(function(x){ var t=new Date(String(x.ts).replace(' ','T')); return t>=d7; }).length;
  var top = Object.keys(perUser).sort(function(a,b){return perUser[b]-perUser[a];}).slice(0,3);
  var prods = Object.keys(prod).sort(function(a,b){return prod[b]-prod[a];}).slice(0,4);
  var aCnt = R.filter(function(r){return r.grade==='A' && latest[r.n];}).length;
  var txt = '📣 裕安11月专线走访战报\\n'
   + '回填进度：'+done+'/'+total+'（'+(total?Math.round(done/total*100):0)+'%）· 近7天新增'+week+'条\\n'
   + '结果：成交'+cnt.A+' · 高意向'+cnt.B+' · 观望'+cnt.C+' · 无意向'+cnt.D+' · 未接触'+cnt.E+'\\n'
   + 'A级重点已回填：'+aCnt+'/22\\n'
   + '行动TOP：'+(top.join('、')||'暂无')+'\\n'
   + '意向产品TOP：'+(prods.join('、')||'暂无')+'\\n'
   + '——大家继续冲，A级客户两周内首轮上门';
  window._rpt = txt;
  document.getElementById('mBody').innerHTML = '<h3>📣 战报（复制后发群）</h3><div class="minfo" style="white-space:pre-wrap;font-size:12.5px">'+txt+'</div><button class="btnGo" onclick="copyReport()">📋 复制战报</button>';
  document.getElementById('ov').classList.add('on');
}
function copyReport(){
  var ta=document.createElement('textarea'); ta.value=window._rpt; document.body.appendChild(ta); ta.select();
  try{ document.execCommand('copy'); toast('战报已复制，去群里粘贴'); }catch(e){ toast('复制失败，请手动选择'); }
  document.body.removeChild(ta);
}
function closeM(){ document.getElementById('ov').classList.remove('on'); }
function loadResults(){
  fetch('/api/results').then(function(x){return x.json();}).then(function(rs){
    RS = rs||[]; latest = {};
    RS.forEach(function(x){ latest[x.name]=x; });
    renderHead(); renderUsers(); chips(); render(); renderGantt(); renderOpp();
  }).catch(function(){ renderHead(); renderUsers(); chips(); render(); renderGantt(); renderOpp(); });
}
document.getElementById('q').addEventListener('keydown', function(e){ if(e.key==='Enter'){ Q=this.value.trim(); render(); } });

// 初始化：凭令牌加载数据
(function init(){
  if(!MYTOK){ showGate(); return; }
  DATAURL += (QS.get('k') ? 'k=' : 'u=') + encodeURIComponent(MYTOK);
  fetch(DATAURL).then(function(x){ return x.json(); }).then(function(d){
    if(!d.ok){ showGate(); return; }
    IS_ADMIN = !!d.admin;
    USER = d.user || '';
    R = d.recs; USERS = d.users || [];
    GENDATE = d.genDate || '';
    EXPURL = '/export.xlsx?k=' + encodeURIComponent(MYTOK);
    document.getElementById('app').style.display='block';
    loadResults();
  }).catch(function(){ showGate(); });
})();
</script>
</body></html>"""
SALES_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>裕安 · 上门营销话术速查</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Microsoft YaHei',sans-serif;background:#eef2f7;color:#333;line-height:1.65;font-size:14.5px;padding:12px}
.hd{background:linear-gradient(135deg,#0a4da3,#1c86e8);color:#fff;border-radius:12px;padding:15px 18px;margin-bottom:12px}
.hd h1{font-size:18px}
.hd p{font-size:11.5px;opacity:.9;margin-top:4px}
.sec{background:#fff;border-radius:11px;margin-bottom:10px;overflow:hidden;box-shadow:0 1px 4px rgba(10,77,163,.08)}
.sec h2{background:#0a4da3;color:#fff;font-size:14.5px;padding:10px 14px;cursor:pointer;display:flex;justify-content:space-between;align-items:center}
.sec h2 .arr{transition:transform .2s}
.sec.open h2 .arr{transform:rotate(90deg)}
.sec .body{display:none;padding:12px 14px}
.sec.open .body{display:block}
.tag{color:#b96a00;font-weight:700}
.q{color:#0a4da3;font-weight:700}
.a{color:#333}
li{margin:7px 0;list-style:none;padding-left:2px}
li b{color:#0a4da3}
.hr{border-top:1px dashed #cfe0f0;margin:9px 0}
.ft{text-align:center;color:#8898aa;font-size:11px;margin:12px 0 20px}
</style></head><body>
<div class="hd"><h1>📖 上门营销话术速查</h1><p>跨境84家+他网12家 · 上门/电话前点开 · 配合走访登记系统使用</p></div>

<div class="sec open"><h2>📌 上门前：三查一带 <span class="arr">▸</span></h2><div class="body">
<li><b>三查</b>：①来源与现状（跨境=主攻专线；他网=官网在友商机房）②风险等级（<span class="tag">需关注</span>=春秋/宏科特/汇德镁，合同从严）③关键人电话</li>
<li><b>一带</b>：测速工具、受理终端、火情AI单页、名片</li>
<li><b>同行楼栋</b>：同楼客户按“存量→A级→B级”排，先谈成一户再同楼复制案例</li>
</div></div>

<div class="sec"><h2>🗣 开场白三套 <span class="arr">▸</span></h2><div class="body">
<li><span class="q">存量客户：</span>“X总，我是电信裕安的XX，今天上门一是给您网络<b>免费体检</b>，二是看您业务到期/提速有没有优惠提前安排，顺便看办公网、收银用得顺不顺。”（服务姿态进门）</li>
<li><span class="q">新客（跨境）：</span>“X总，我是属地电信裕安的XX，看到咱们在<b>海关有进出口备案</b>，平时接触不少跨境企业，今天路过来认识一下，帮您看看现在网络够不够用——不办业务没关系，留个联系方式，以后网络有问题找我最快。”</li>
<li><span class="q">他网客户：</span>“X总，我们在做辖区企业官网访问体检，发现咱们官网还有提升空间——我现场测给您看（<b>当场测速对比</b>），今天反馈下情况，聊聊有没有优化方案。”</li>
</div></div>

<div class="sec"><h2>🌍 跨境客户五场景 <span class="arr">▸</span></h2><div class="body">
<li><span class="q">跨境专线（主攻70家空白）：</span>“做跨境的，网络就是门面。普通宽带走公众国际出口，高峰挤；<b>跨境专线独立通道、直连国际优质出口，稳不抖不丢包</b>。您现在专线是空白，首装有优惠。”（先问：“高峰卡不卡？丢过单没有？”再报价）</li>
<li><span class="q">SD-WAN：</span>“多个店铺/仓库想连成一张网，一条条拉专线周期长成本高；<b>SD-WAN智能选路、开通快、按需扩</b>。”</li>
<li><span class="q">固定IP：</span>“出差想连回公司系统连不上？<b>固定IP=公司网络的固定门牌号</b>，远程访问、报关对接、监控回传都靠它。”</li>
<li><span class="q">商企宽带：</span>“您现在用的家用套餐是高峰共享；<b>商企=独享带宽+固定IP+装维随叫随到</b>，收银/打印机/监控互不干扰。”</li>
<li><span class="q">存量升档：</span>“看您光猫/设备是旧版，<b>型号差一代、带机量和稳定性差很多</b>（拿图片对比），到期正好升档换新设备，不额外花多少。”</li>
</div></div>

<div class="sec"><h2>🔧 他网策反两套</h2><div class="body">
<li><span class="q">自有服务器（A记录直连）：</span>“刚才测速看到了，官网现在在联通/移动机房。服务器在哪，客户访问的“第一公里”就在哪。<b>电信可以给您“服务器迁移+专线+固定IP”一揽子</b>，设备是您的，只是搬家不是重买。您机房合同什么时候到期？到期前我们出对比方案，<b>您还能拿我们方案压现在价格</b>。”</li>
<li><span class="q">建站托管（CNAME）：</span>“您官网是建站公司代管吧？每年续费多少？<b>天翼云建站，机房一级骨干节点、价格公开</b>，迁过来先跑着对比，不合适随时迁回——域名和您的。”</li>
</div></div>

<div class="sec"><h2>✅ 现场必做五动作</h2><div class="body">
<li><b>1 测速留痕</b>：当场测速截图，下次上门的证据</li>
<li><b>2 拍照登记</b>：光猫/路由型号、布线（水印相机）</li>
<li><b>3 火情AI必提</b>：“厂房/仓库消防安全，我们有个AI预警，符合条件基本免费装，明天师傅顺带办了”（三倍激励60元/路）</li>
<li><b>4 登记到期时间</b>：现有宽带哪家、多少钱、<b>什么时候到期</b>=策反时间表</li>
<li><b>5 承诺闭环</b>：离场明确“明天X点发您方案/回电话”，说时间不说“回头联系”</li>
</div></div>

<div class="sec"><h2>🛡 异议六连答</h2><div class="body">
<li><span class="q">“现在用的挺好”：</span>“留我个电话，<b>网络这东西不敢保证天天正常，下次有问题两小时到</b>，存个备用工程师。”</li>
<li><span class="q">“比友商贵”：</span>“一分钱一分货，专线是独享通道他网多是共享带宽。我出两家<b>实测对比</b>给您，不比广告比实测。”</li>
<li><span class="q">“要商量”：</span>“应该的，方案+测速图整理一页发您，明天上午电话对一下。”</li>
<li><span class="q">“要搬走/缩减”：</span>“号码宽带都能<b>平移</b>，搬家前说一声就行；真要缩减我看怎么处理最省钱。”（先服务防拆机）</li>
<li><span class="q">“你们服务不行”：</span>“您说的我记下了，我就是这片负责人，今天起这事找我。”（不辩解，接管）</li>
<li><span class="q">“再等等”：</span>“行，唯一别的是别到期前被他网低价绑两年——您把到期日告诉我，到期前一个月我给您个方案，主动权在您。”</li>
</div></div>

<div class="sec"><h2>🚪 离场与跟进</h2><div class="body">
<li>当场加微信，发“X总我是电信裕安XX”</li>
<li><b>24h内</b>发一页方案（测速对比+报价+同楼案例）</li>
<li><b>3天内</b>电话回访；未成交按到期时间进台账，<b>到期前1个月黄金跟进</b></li>
<li><b>同楼案例最强话术</b>：“您隔壁XX公司上个月刚办”</li>
</div></div>

<div class="sec"><h2>⚠️ 风险纪律</h2><div class="body">
<li>3家<span class="tag">需关注</span>（春秋电子/宏科特/汇德镁）：合同分期从严、别送设备</li>
<li>不贬友商、不虚承诺（承诺了就兑现，晨会回查）</li>
<li>登记走访结果当天回填（每人专属链接）</li>
</div></div>

<div class="ft">裕安政商 · 走访营销配套 · 本地离线运行 · 与话术手册Word版一致</div>
<script>
document.querySelectorAll('.sec h2').forEach(function(h){ h.onclick=function(){ h.parentNode.classList.toggle('open'); }; });
</script>
</body></html>"""

def load_results():
    if os.path.exists(RESULTS):
        try:
            with open(RESULTS, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_results(rs):
    with open(RESULTS, 'w', encoding='utf-8') as f:
        json.dump(rs, f, ensure_ascii=False, indent=1)

class H(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=BASE, **kw)
    def log_message(self, *a):
        pass
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, must-revalidate')
        super().end_headers()
    def _json(self, obj, code=200):
        body = obj.encode('utf-8') if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)
    def do_GET(self):
        qs = urllib.parse.urlparse(self.path)
        p = qs.path
        q = urllib.parse.parse_qs(qs.query)
        if p in ('/', '/index.html'):
            body = VISIT_HTML.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif p == '/visit.html':
            body = VISIT_HTML.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif p == '/sales.html':
            body = SALES_HTML.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif p == '/api/results':
            with LOCK:
                rs = load_results()
            self._json(rs)
        elif p == '/api/data':
            k = (q.get('k') or [''])[0]
            u = (q.get('u') or [''])[0]
            if k and k == TOK_ADMIN:
                self._json({'ok': True, 'admin': True, 'recs': RECS, 'users': USERS,
                            'genDate': datetime.date.today().strftime('%Y-%m-%d')})
            elif u and u in TOK_USERS:
                name = TOK_USERS[u]
                mine = [x for x in RECS if x['assign'] == name]
                self._json({'ok': True, 'user': name, 'recs': mine,
                            'genDate': datetime.date.today().strftime('%Y-%m-%d')})
            else:
                self._json({'ok': False}, 403)
        elif p == '/export.xlsx':
            k = (q.get('k') or [''])[0]
            if k != TOK_ADMIN:
                self._json({'ok': False, 'err': 'unauthorized'}, 403)
                return
            try:
                body = self._build_export()
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="visit_export.xlsx"')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self._json({'ok': False, 'err': str(e)[:100]}, 500)
        else:
            self._json({'ok': False, 'err': 'not found'}, 404)
    def do_POST(self):
        p = urllib.parse.urlparse(self.path).path
        if p == '/api/submit':
            try:
                n = int(self.headers.get('Content-Length', 0) or 0)
                data = json.loads(self.rfile.read(n).decode('utf-8'))
                u = str(data.get('tok', ''))
                user = TOK_USERS.get(u, '')
                if not user:
                    self._json({'ok': False, 'err': 'unauthorized'}, 403)
                    return
                rec = {
                    'name': str(data.get('name', ''))[:80], 'user': user,
                    'date': str(data.get('date', ''))[:12], 'result': str(data.get('result', ''))[:4],
                    'products': [str(x)[:12] for x in (data.get('products') or [])][:10],
                    'note': str(data.get('note', ''))[:400], 'expire': str(data.get('expire', ''))[:14],
                    'next': str(data.get('next', ''))[:14],
                    'ts': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                }
                if not rec['name']:
                    self._json({'ok': False, 'err': 'empty'}, 400)
                    return
                with LOCK:
                    rs = load_results()
                    rs.append(rec)
                    save_results(rs)
                self._json({'ok': True, 'count': len(rs)})
            except Exception as e:
                self._json({'ok': False, 'err': str(e)[:80]}, 500)
        else:
            self._json({'ok': False, 'err': 'not found'}, 404)
    def _build_export(self):
        rs = load_results()
        latest = {}
        for r in rs:
            latest[r['name']] = r
        RT = {'A': 'A-已成交', 'B': 'B-高意向', 'C': 'C-观望', 'D': 'D-无意向', 'E': 'E-未接触上'}
        wb = Workbook()
        ws = wb.active
        ws.title = '走访回填'
        cols = [('清单来源',14),('优先级',12),('分配人',9),('企业名称',26),('法定代表人',9),('联系电话',15),
                ('命中网格',13),('注册地址',30),('企业规模',9),('参保人数',8),('行业大类',12),('风险等级',9),
                ('产品建议',28),('走访日期',11),('走访结果',14),('意向产品',24),('客户到期时间',12),('下次跟进',12),('备注',26),('登记人',9),('登记时间',16)]
        BLUE='FF0A4DA3'
        thin=Side(style='thin',color='FFD5DFEA')
        border=Border(left=thin,right=thin,top=thin,bottom=thin)
        for ci,(h,w) in enumerate(cols,1):
            c=ws.cell(row=1,column=ci,value=h)
            c.font=Font(bold=True,color='FFFFFFFF',size=10.5,name='微软雅黑')
            c.fill=PatternFill('solid',start_color=BLUE)
            c.alignment=Alignment(horizontal='center',vertical='center')
            c.border=border
            ws.column_dimensions[get_column_letter(ci)].width=w
        ws.row_dimensions[1].height=26
        ws.freeze_panes='D2'
        for ri,r in enumerate(sorted(RECS,key=lambda x:(x['assign'],x['n'])),2):
            l=latest.get(r['n'])
            vals=[r['src'],r['grade'],r['assign'],r['n'],r['legal'],r['tel'],r['grid'],r['addr'],r['size'],
                  r['staff'],r['ind'],r['risk'],r['rec'],
                  l['date'] if l else '',RT.get(l['result'],'') if l else '',
                  '、'.join(l['products']) if l else '',l['expire'] if l else '',
                  l['next'] if l else '',l['note'] if l else '',l['user'] if l else '',l['ts'] if l else '']
            for ci,v in enumerate(vals,1):
                c=ws.cell(row=ri,column=ci,value=v)
                c.font=Font(size=10,name='微软雅黑')
                c.border=border
                c.alignment=Alignment(vertical='center',horizontal=('left' if ci in(4,8,13,19) else 'center'),wrap_text=(ci in(8,13,19)))
                if l and ci==15:
                    color='FF1E8449' if l['result'] in('A','B') else('FFB7950B' if l['result']=='C' else 'FF999999')
                    c.font=Font(size=10,name='微软雅黑',bold=True,color=color)
                elif not l and ci==15:
                    c.font=Font(size=10,name='微软雅黑',color='FFC0392B',bold=True)
                    c.value='待走访'
                if ri%2==0:
                    c.fill=PatternFill('solid',start_color='FFEAF3FB')
        ws.auto_filter.ref='A1:%s%d'%(get_column_letter(len(cols)),len(RECS)+1)
        buf=io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

if __name__ == '__main__':
    print('Yunan Visit Cloud Server starting on 0.0.0.0:80')
    server = ThreadingHTTPServer(('0.0.0.0', 80), H)
    server.serve_forever()
