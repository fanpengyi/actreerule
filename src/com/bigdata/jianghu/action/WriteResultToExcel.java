package com.bigdata.jianghu.action;

import com.bigdata.jianghu.utils.ArticleBean;
import com.bigdata.jianghu.utils.ExcelTest;
import com.bigdata.jianghu.utils.ExportExcel2007;
import com.facebook.presto.jdbc.internal.guava.collect.Sets;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WriteResultToExcel {



	static List<String> regs = new ArrayList<String>();
	static Set<String> newReg = new HashSet<String>();

	public static void init(){

		regs.add("(.*(宕机|断网|电力中断|制冷中断|网络中断|非正常(停|死)机|异常(停|死)机|云计算服务中断|延误|航班取消|航班返航|迫降|备降|bug|系统(挂|崩)了|报错|泄漏(隐私|消息|信息|数据)|(隐私|消息|信息|数据)泄露|(无法|不能)(值机|选座)|数据错误|信息不准|罢工|停运|暴徒|袭警|暴乱|系统瘫痪|系统故障|系统问题|网络攻击|反垄断法|通用数据保护条例|PNR抓取|票面信息提取|PID采屏|旅客隐私数据|灾备|容灾|停机|故障|事故|旅客堆积|旅客滞留|投诉|数据丢失|安全隐患|差评|侵权|漏洞|攻击|起诉|违法|违规|侵害|侵犯|警报|被停飞|超卖|(吃|要)回扣|\\\\d+(人)?(死|亡)|\\\\d+死\\\\d+伤|死了\\\\d+人|(乱|被|高)收费|影响(不好|很糟|糟糕|出行)|电梯(突然)?(逆行|倒转)|售票机前(比较|很|非常|十分|特别)?拥挤|(无|没)人(急救|救命|施救|相救|敢救)|(发生|闹)口角|妨(碍|害)公务|(爆|暴|曝)粗口|出了?(问题|事).{0,3}没有监控|违反.{0,3}法|态度.{0,3}差|滥用.{0,4}职权|扇.{0,4}耳光|掩盖(责任|事实|真相)|超售(现象)?严重|(捆绑|隐性)消费|重大(交通|安全)?事故|工资(太?)低|(太|真|很)(烂|差)|劳(动|务)(纠纷|争议)|效率(低|太?慢)|拉客仔|发动机故障|迫降|黑的士|喊口号|遭追打|被拘|互殴|炸飞机|强开舱门|野蛮|漏雨|航班取消|航班延误|扰乱机场秩序|滞留|情绪焦躁|不理智|纠纷|打砸|抗议|处置不力|通报|批评|空闹|对骂|行李箱刮坏|暴力搬运|航延|瘫痪|态度差|情绪激动|空喊口号|大闹|涉嫌|妨碍|刑拘|气到爆炸|恶心|恶劣|(造成|导致).{0,4}(误机|延误)|骂声|强行|吓人|值机.{0,3}乱|垃圾|航班误点|效率更?低|吐槽|严重超售|没有天理|重大(交通)?事故|爆炸|暴利|提起公诉|索赔|不作为|瞒报|漏报|延报|内幕|黑幕|失事|违规|违法|非法|不合理|刮蹭|侵权|赔偿|殴打|辱骂|侮辱|闹事|阻挠|打人|冲突|谩骂|存疑|质疑|疑似|疑为|遭疑|遭批|被批|被疑|爆料|被爆|被曝|惊曝|频曝|媒曝|媒体曝|网曝|曝光|被指|被调查|被审|被罚|陷入|深陷|又陷|再陷|屡陷|惊现|频发|回应|澄清|揭露|受困|困局|乱局|乱象|隐忧|堪忧|非议|风波|事发|案发|被诉|上诉|起诉|举报|投诉|诟病|丑闻|获赔|涉事|责令|引热议|窝案|引发|争议|叫停|痛批|挪用|玩忽职守|纵容|包庇|被撤|撤职|造假|作假|虚假|做假|失火|大火|起火|讹诈|泄漏|死亡|撞亡|失踪|失联|被困|逃生|遇难|经济问题|坑爹|焦糊味|通奸|闹剧|砸坏|砸伤|猛泼|被砸|暴力|腐败|受贿|索贿|收贿|贿赂|违纪|贪腐|贪污|反腐|免职|落马|双规|双开|停职|强制|信息售卖|包养|包二奶|盗刷|盗用|套现|诉至|败诉|反诉|诉讼|开庭|审理|被告|原告|诉称|被免|去职|调离|行贿|涉贪|工资太?低|降薪|黑心|太黑|剥削|不平等|歧视|尼玛|排长队|乱收费|被查|频现|敛财|隐患|炮轰|勒令|欺诈|逃税|罚款|推卸责任|变相|处罚|无耻|卷款|巨贪|严重违纪|跑路|巨额|高利贷|涉案|涉腐|贿款|信息泄漏|暴利非法|揭发|检举|暗箱操作|诈骗|霸王|驳斥|立案|阴谋|万恶|亏损|巨亏|信用证案|潜逃|外逃|公款|革职|盛传|钓鱼|假冒|冒充|停摆|失望|系统故障|虚高|误导|欺骗|枉费|伤心|乌龙|伤不起|倒霉|糟糕|无语|折腾|苦逼|坑人|坑死|悲催|讨厌|破银行|倒闭|办不了|重复计费|离谱|等很久|不尊重|闲着|太牛|扯蛋|假钞|傻逼|落伍|借口|我擦|脑残|超级慢|急疯|焦虑|垄断|鄙视|耗着|神马玩意|恶性竞争|绝望|陷阱|乱扣|冤枉|滞后|霸王条款|立案调查|信息泄露|急速扩张|理财亏损|携巨款|挪用公款|系统停摆|卷入|变脸|疑点|网友曝|被判|被盗|打劫|泄露|经济犯罪|逮捕|压榨|债券罪|接受调查|骗取|贪官|刑罚|刑责|嫌疑|获刑|领刑|判刑|窝点|身亡|专案|获罪|渎职|坠亡|抢劫|盗得|盗窃|失窃|窃贼|持刀|持枪|杀人|被杀|拒不|盗取|击毙|拖欠|遭受|被控|勾结|以权谋私|查获|行骗|骗局|侵占|工亡|恶意|不诚信|可恶|遭抢|自杀|坏账|反价|被抢|追讨|侵吞|违约|冒名|伪造|情妇|案件|洗黑?钱|被贷款|高危漏洞|失控|推搡|撕扯|猥琐|动手|叫嚣|操作失误|偷税|漏税|推卸|推诿|飞单|被传|牟取|谋取|利益输送|被拒|被搜身|被掀|不翼而飞|冒领|不知情|冲撞|反抗|扣压|闹僵|蹊跷|无果|恶性|罚单|太垃圾|罗生门|骗子|扰民|讨说法|偷运|暗中操控|败露|报复|爆粗口|被霸占|被逼|被篡改|被盗用|被扣|被拦|被起诉|被欠|被烧|被收|被占|不干人事|不合法|不合规|不堪入耳|不守规矩|不以为耻|惨重|藏匿|查出|超期|扯皮|趁火打劫|臭不可闻|出卖|处理不当|畜牲|触目惊心|串通|吹嘘|篡改|擦边球|胆大妄为|道德败坏|低劣|低俗|抵赖|颠倒黑白|垫底|顶撞|动粗|断裂|恶炒|恶果|恶俗|公愤|狗腿子|故意为难|害苦|黑钱|黑手|哄骗|胡搅蛮缠|糊弄|唬人|荒唐|或涉|祸害|讥讽|假借|见钱眼开|狡辩|叫骂|纠集|居心不良|居心何在|居心叵测|拒之门外|聚众哄抢|可疑|恐慌|捆绑|拉黑|来路不明|赖帐|赖账|老赖|冷嘲热讽|利益纠葛|良心泯灭|劣迹|令人作呕|漏洞百出|沦丧|屡禁不止|骂人|埋怨|蛮不讲理|蛮横专断|瞒天过海|满嘴大话|冒名顶替|没兑现|昧良心|昧着良心|蒙混|明目张胆|抹黑|难辞其咎|难闻|捏造|骗招|频涉|欺瞒|欺辱|牵涉|强烈不满|强令|抢夺|缺德|缺乏人性|惹不满|惹官司|惹祸|人心惶惶|丧尽天良|擅改|擅自|伤天害理|上当|涉瞒|视而不见|受罚|数字不实|双重标准|私改|私通私占|私自|搜刮|损毁|损人利己|索偿|贪财|逃避|痛斥|痛苦不堪|痛述|推托|推脱|歪理|歪门邪道|威逼|未经同意|龌龊|龌蹉|污秽|污蔑|诬陷|无德|无恶不作|无人接听|误扣|下流|消息不实|心黑|信任危机|形同虚设|血本无归|严重影响|掩耳盗铃|一拖再拖|疑团|疑遭|引不满|引争议|炸弹|诱导|预谋|欲盖弥彰|怨言|栽赃|遭毁|遭拒|造孽|造谣|责罚|责骂|张冠李戴|争吵|指控|致命|致伤|混乱|置之不理|失职|踢皮球|难办理|无人管|没人管|去向不明|降职|截留|恐吓|被稽查|成摆设|挥霍|据为己有|袒护|无人过问|喊冤|抓捕归案|被批捕|被处理|被带走|大摆筵席|大摆宴席|索拿卡要|天价酒|天价烟|通风报信|严重警告|赃款|置若罔闻|众怒|私征|阻扰|纂改|罪恶|作恶|被追责|强毁|被毁|被骗|被偷|不公|不满|惨剧|超标|吊销|出事|冒用|发指|封口费|坑害|编造|告发|频生|强扣|逼迫|枉法|受害|遭调查|端掉|打掉|破获|争执|诽谤|不正当|剽窃|谴责|蛀虫|打击报复|败类|死刑|苦不堪言|殃及|猫腻|声讨|谎称|破损|寻衅滋事|偷窥|牟利|变造|欺压|错案|冤案|恶行|维权无门|凌辱|欺凌|责成|擅自撤离|假充|假装|充作|混充|强迫|作奸犯科|为非作歹|专横跋扈|作威作福|肆无忌惮|偷窃|争论|咒骂|诅咒|指责|斥责|训斥|伙同|卑劣|刁难|陷害|棍骗|诈欺|诱骗|诳骗|人渣|畜生|苦不可言|里应外合|马脚|千疮百孔|逃跑|挑事|欺侮|滋事|肇事|罪行|套用|欠贷|断贷|假币|空头账户|乱拆借|骗存|吞钱|无法兑付|无法提现|无故冻结|吸存|限制提现|以贷谋私|偷刷|印子钱|暗藏|暗地|暗斗|暗访|暗算|暗箱|肮脏|把柄|霸占|报料|报应|抱怨|暴打|暴光|暴露|卑鄙|悲剧|被捕|被打|被夺|被黑|被禁|被坑|被喷|被迫|被删|逼宫|逼死|闭门羹|庇护|弊病|剥夺|不符|不公道|不力|不落实|不忍直视|不知羞耻|艹|查办|嘲弄|惩罚|仇恨|臭名远扬|出尔反尔|吹捧|粗暴|粗口|粗鲁|猝死|存忧|打假|打伤|大案|大骂|大失所望|道歉|诋毁|抵制|丢脸|丢人|毒瘤|对峙|恶毒|恶习|发飙|发火|贩假|仿冒|仿造|风口浪尖|封锁|敷衍|搞垮|告破|公然|攻击|勾当|鼓吹|故伎重演|挂掉|关停|官商乡黑|官商作风|官司|冠冕堂皇|过失|黑名单|忽悠|狐狸尾巴|怀疑|荒谬|灰幕|毁掉|混蛋|激怒|伎俩|架空|狡诈|搅局|搅乱|揭穿|进水|警告|纠缠|拒赔|开骂|抗争|控诉|口角|口水战|口水仗|扣留|哭天喊地|滥用|狼狈为奸|劳民伤财|雷人|令人担忧|流氓|屡遭|买通|蛮横|矛盾|冒牌|没人性|蒙蔽|蒙骗|蒙人|灭亡|明争暗抢|抹杀|幕后操纵|恼火|内斗|内讧|内奸|你妈|扭曲|弄虚作假|怒斥|拍马屁|赔光|喷子|抨击|批驳|批判|辟谣|骗人|撇清|频出|破绽|曝出|曝料|欺负|企图|牵扯|牵连|潜规则|遣责|欠薪|强加|强压|窃取|侵犯|侵害|侵扰|禽兽|取缔|扰乱|骚扰|杀害|煞笔|伤亡|深恶痛绝|深受其害|失实|失误|示威|受骗|受阻|耍赖|衰落|私通|私占|撕毁|死伤|肆意|耸人听闻|搜身|损害|套取|替死鬼|替罪羊|天理何在|挑拨|挑衅|跳楼|偷换|偷拍|图谋不轨|托词|唾弃|歪曲|网传|危言耸听|违背|违反|违禁|围攻|未兑现|诬告|无赖|无良|误操作|误判|误事|袭击|洗脑|下三滥|下作|虚报|血汗钱|压制|延误|严厉谴责|言而无信|掩盖|妖言惑众|谣言|疑是|阴险|隐瞒|营私舞弊|欲诉|遇阻|暂扣|遭到|遭否|遭损|糟蹋|诈捐|质问|作死|状告|追究|追责|走过场|走后门|遭罚|拎包贼|小偷|猖獗|扣押|捧臭脚|跪舔|懒政|晚点|闷热|撞人|车祸|丧命|火灾|撞机|事故|噩梦|密集恐惧症|要人?命|密恐症|被宰|砍人|劫杀|消极怠工|超负荷|怒怼|热晕|空调不开|不开空调|黑飞|走私|恐怖|撞坏|紧急(着陆|备降)|不慎|跌落|骨折|昏迷|昏厥|摔下|摔倒|浩劫|违法?停|打压|漏水|暴跌|滔天大祸|负面影响|亏钱|对人体有害|伸冤无门|重伤|涉黑|噪音|恶性公关|幺蛾子|遭殃|撒泼|负面(新闻)?缠身|刺伤|打斗|不治身亡|甩客|不到位|形象大打折扣|夸大宣传|下降率太高|掌掴|当众大?小便|当众大小?便|不予处理|法院判决|盈利能力最?差|婊子|脏乱差|乱停|大闹[\\\\u3E00-\\\\u9FA5]*机场|利用职务之便|尾随|严肃处置|存在安全(因素|问题)|打脸|望而生畏|重复安检|手续繁杂|交通复杂|信号不稳|信号差|信号不好|信号烂|常年不稳定|踢|工人工资不给|工人急等钱|等着血汗钱|让人头疼|没有拿到补偿款|压到现在不给|没地说理|净空保护区范围|噪音区安置房|拖欠工资|劣质水泥|欠工资不还|建设问题|着火|冒烟|超售|不明飞行物|ufo|中国为什么行|WHY .?CHINA|why.?china|没有结清|(演练|测试|模拟)(失败|失误|错误|不成功|没有成功|砸|出现问题|中断|站厅|中止|拉条幅|拉横幅|上访)|拆迁黑慕|补偿过少|闹哪样|错误|不顾|失败|餐馆少|喧闹|不便|尚未启用|走错机场|飞机故障|打车难|飞机事故|掉价|配套(设施)?不全|廊桥故障|黑出租|装修味道|刺鼻|摆渡车拥堵|交通拥堵|故障|心情.{0,3}(不好|差)|公器私用|(素质|态度|水平).{0,3}(差|不好|恶略)|低下|低端|低档次|太垃圾|很差|太弱|一落千丈|荡然无存|追星|戏谑|给个说法|错了不道歉|公号私用|毫无信服力|搞臭|鸡你太美|误操作|不专业|骂官微|低级|能力差|不负责|不规范|无改错.{0,3}心|没有投诉途径|监督无举报|一再犯错|不知悔改|靠关系|不当一回事|搞糊|lou|LOU|Lou|low|LOW|Low|喽|差评|不满意|很坏|尴尬|航班.{0,4}少|太贵|走错|便捷性.{0,3}(不好|差|不如)|甲醛味|冤枉路|等了很久|破飞机|烦死|有什么好|不方便|交通不便|严重交通问题|衰到极致|太吵了|坏了|运能过剩|极性.{0,2}低|好着急|三无产品|空标建筑|坑死|没走不让我?上|一点都不好|太远了|打扰.{0,3}休息|差点意思|没有操守|很懵|混乱|不如首都机场|傻.?大|炸毛|太呛|有待完善|稍显不足|有点远|一手遮天|一言堂|神翻译|脸都丢尽|丢脸|(翻译|拼写|书写|标志|标示|导航)(错误|不清楚|不清晰|无法辨别|模糊)|错误(翻译|拼写|书写|标志|标示))){6,}.*");


		try {
			//BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(new File("D:\\work\\Linux\\match\\ceshi\\reg1069")),
			BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(new File("./resources/regex1069")),
					"UTF-8"));
			String lineTxt = null;
			while ((lineTxt = br.readLine()) != null) {
				String[] rulecontent=lineTxt.split(" ");
				newReg.add(rulecontent[0]);
			}
			br.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}


	public static void main(String[] args) throws ParseException, IOException, SQLException {


		List<ArticleBean> inputList = new ArrayList<ArticleBean>();
		ExcelTest excelTest=new ExcelTest();
		//Workbook wb = excelTest.getExcel("D:\\work\\Linux\\test\\input.xlsx");
		Workbook wb = excelTest.getExcel("./resources/input.xlsx");
		if(wb==null)
			System.out.println("文件读入出错");
		else {
			inputList=excelTest.analyzeExcel(wb);
		}

		//System.out.println(list);

		//使用老程序写出耗时

		//inputList = inputList.subList(0,100);

		init();


		List<Map<String, String>> resultList = new ArrayList<>();


		Map<String, Map<String,String>> resultMap = new HashMap<>();

		getOldRuleResult(inputList,resultMap);


		getNewRuleResult(inputList,resultMap);


		for (String key : resultMap.keySet()) {
			Map<String, String> map = resultMap.get(key);
			String oldResult = map.get("result");
			String newResult = map.get("newResult");
			if(StringUtils.isNotBlank(newResult)){
				if(newResult.equals(oldResult)){
					map.put("old2newCompare","true");
				}else{
					map.put("old2newCompare","false");
				}

			}
			resultList.add(map);
		}


		String filename =  "./resources/1069-result.xlsx";

		XSSFWorkbook wb2 = new XSSFWorkbook();
		ExportExcel2007 e = new ExportExcel2007(filename, wb2);
		XSSFSheet sheet1 = wb2.createSheet();

		String[] columns = new String[]{"id","content","costTime","result","newCostTime","newResult","newtotlecount","old2newCompare"};

		e.createExcel(sheet1, resultList,columns ,
				columns);
		e.outToFile();

	}

	private static void getOldRuleResult(List<ArticleBean> list, Map<String, Map<String,String>> resultMap) {
		try {
			for(ArticleBean article :list){
				HashMap<String, String> map = new HashMap<>();
				long a= System.currentTimeMillis();//获取当前系统时间(毫秒)

				boolean z= ExcelTest.isContainChinese(article.getContent(),regs.get(0));

				long costTime = System.currentTimeMillis() - a;

				map.put("id",String.valueOf(article.getId()));
				map.put("content",article.getContent());
				map.put("costTime",String.valueOf(costTime));
				if(z){
					map.put("result","success");
					System.out.println("成功");
				}
				resultMap.put(String.valueOf(article.getId()),map);
				System.out.println(resultMap.size());
			}

		} catch (Exception e) {
			System.err.println("read errors :" + e);
		}
	}



	private static void getNewRuleResult(List<ArticleBean> inputList, Map<String, Map<String,String>> resultMap) {
		try {

			ACTrie trie = new ACTrie();
			try {
				//BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(new File("D:\\work\\Linux\\match\\ceshi\\word1069")),

				BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(new File("./resources/1069/1069_words")),
						"UTF-8"));
				String lineTxt = null;
				HashSet<Object> set = Sets.newHashSet();
				while ((lineTxt = br.readLine()) != null) {
					String[] rulecontent=lineTxt.split(" ");
					set.add(rulecontent[0]);
//                System.out.println(lineTxt);
				}
				for (Object o : set) {
					trie.addKeyword((String)o);
				}
				br.close();
			} catch (Exception e) {
					e.printStackTrace();
			}

			String lineTxt = null;
			BufferedReader br = null;
			try {
				for (ArticleBean articleBean : inputList) {

					int totlecount=0;
					Map<String, String> oldMap = resultMap.get(String.valueOf(articleBean.getId()));

					long a= System.currentTimeMillis();//获取当前系统时间(毫秒)

					Collection<ACTrie.Emit> emits = trie.parseText(articleBean.getContent());
					for (ACTrie.Emit emit : emits) {
						totlecount++;
					}


					for(String key :newReg){
						Pattern p = Pattern.compile(key);
						Matcher m = p.matcher(articleBean.getContent());
						int count = 0;
						while (m.find()) {
							totlecount++;
							count++;
						}
						//System.out.println(key+":"+count);
					}


					if(totlecount >= 6){
						oldMap.put("newResult","success");
					}
					long newCostTime = System.currentTimeMillis() - a;
					System.out.println("costTime:"+newCostTime);

					oldMap.put("newCostTime",String.valueOf(newCostTime));
					oldMap.put("newtotlecount",String.valueOf(totlecount));
				}

//            System.out.println("a : " + Collections.frequency(list, "很差"));
			} catch (Exception e) {
				e.printStackTrace();
			}

		} catch (Exception e) {
			System.err.println("read errors :" + e);
		}
	}




}
