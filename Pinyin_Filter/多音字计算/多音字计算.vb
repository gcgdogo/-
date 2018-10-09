'以33000字典为基础，计算所有多音字
Sub MutiPinyin_Calculate()
	Dim RegExp_Match as new RegExp, Pinyin as String
	Dim RegExp_PyChk as new RegExp, Pinyin_Add as String ,Pinyin_Add_Count as Integer
	Dim Pinyin_Dict as String , DictFinger as Long ,DictFinger_Char as String
	Dim RegExp_MatchCollection as MatchCollection , I as Integer
	Dim Time_ST as Double

	Time_ST=Timer()

	'vvvv好长一段字典啊！！！

	'33000版多音字太多了，看来那些不常用的字拼音更复杂，改用比较准确的7173版字典
		Pinyin_Dict = Pinyin_Dict & "a啊,a阿,a吖,a嗄,a腌,a锕,ai爱,ai矮,ai挨,ai哎,ai碍,ai癌,ai艾,ai唉,ai哀,ai蔼,ai隘,ai埃,ai皑,ai嗌,ai嫒,ai瑷,ai暧,ai捱,ai砹,ai嗳,ai锿,ai霭,an按,an安,an暗,an岸,an俺,an案,an鞍,an氨,an胺,an庵,an揞,an犴,"
		Pinyin_Dict = Pinyin_Dict & "an铵,an桉,an谙,an鹌,an埯,an黯,ang昂,ang肮,ang盎,ao袄,ao凹,ao傲,ao奥,ao熬,ao懊,ao敖,ao翱,ao澳,ao拗,ao媪,ao廒,ao骜,ao嗷,ao坳,ao遨,ao聱,ao螯,ao獒,ao鏊,ao鳌,ao鏖,ao岙,ba把,ba八,ba吧,ba爸,ba拔,ba罢,ba跋,ba巴,"
		Pinyin_Dict = Pinyin_Dict & "ba芭,ba扒,ba坝,ba霸,ba叭,ba靶,ba笆,ba疤,ba耙,ba捌,ba粑,ba茇,ba岜,ba鲅,ba钯,ba魃,ba菝,ba灞,bai百,bai白,bai摆,bai败,bai柏,bai拜,bai佰,bai伯,bai稗,bai捭,bai掰,ban半,ban办,ban班,ban般,ban拌,ban搬,ban版,ban斑,ban板,ban伴,ban扳,"
		Pinyin_Dict = Pinyin_Dict & "ban扮,ban瓣,ban颁,ban绊,ban癍,ban坂,ban钣,ban舨,ban阪,ban瘢,bang帮,bang棒,bang绑,bang磅,bang镑,bang邦,bang榜,bang蚌,bang傍,bang梆,bang膀,bang谤,bang浜,bang蒡,bao包,bao抱,bao报,bao饱,bao保,bao暴,bao薄,bao宝,bao爆,bao剥,bao豹,bao刨,bao雹,bao褒,bao堡,bao苞,"
		Pinyin_Dict = Pinyin_Dict & "bao胞,bao鲍,bao炮,bao龅,bao孢,bao煲,bao褓,bao鸨,bao趵,bao葆,bao勹,bei被,bei北,bei倍,bei杯,bei背,bei悲,bei备,bei碑,bei卑,bei贝,bei辈,bei钡,bei焙,bei狈,bei惫,bei臂,bei褙,bei悖,bei蓓,bei鹎,bei鐾,bei呗,bei邶,bei鞴,bei孛,bei陂,bei碚,bei埤,bei萆,"
		Pinyin_Dict = Pinyin_Dict & "ben本,ben奔,ben苯,ben笨,ben锛,ben贲,ben畚,ben坌,beng蹦,beng绷,beng甭,beng崩,beng迸,beng蚌,beng泵,beng甏,beng嘣,beng堋,bi比,bi笔,bi闭,bi鼻,bi碧,bi必,bi避,bi逼,bi毕,bi臂,bi彼,bi鄙,bi壁,bi蓖,bi币,bi弊,bi辟,bi蔽,bi毙,bi庇,bi敝,bi陛,"
		Pinyin_Dict = Pinyin_Dict & "bi毖,bi痹,bi秘,bi泌,bi秕,bi薜,bi荸,bi芘,bi萆,bi匕,bi裨,bi畀,bi俾,bi嬖,bi狴,bi筚,bi箅,bi篦,bi舭,bi荜,bi襞,bi庳,bi铋,bi跸,bi吡,bi愎,bi贲,bi滗,bi濞,bi璧,bi哔,bi髀,bi弼,bi妣,bi婢,bi埤,bian边,bian变,bian便,bian遍,"
		Pinyin_Dict = Pinyin_Dict & "bian编,bian辩,bian扁,bian贬,bian鞭,bian卞,bian辨,bian辫,bian忭,bian砭,bian匾,bian汴,bian碥,bian蝙,bian褊,bian鳊,bian笾,bian苄,bian窆,bian弁,bian缏,bian煸,biao表,biao标,biao彪,biao膘,biao杓,biao婊,biao飑,biao飙,biao鳔,biao瘭,biao飚,biao镳,biao裱,biao骠,biao镖,biao灬,biao髟,bie别,"
		Pinyin_Dict = Pinyin_Dict & "bie憋,bie鳖,bie瘪,bie蹩,bin宾,bin濒,bin摈,bin彬,bin斌,bin滨,bin豳,bin膑,bin殡,bin缤,bin髌,bin傧,bin槟,bin鬓,bin镔,bin玢,bing并,bing病,bing兵,bing冰,bing丙,bing饼,bing屏,bing秉,bing柄,bing炳,bing摒,bing槟,bing禀,bing邴,bing冫,bo拨,bo波,bo播,bo泊,bo博,"
		Pinyin_Dict = Pinyin_Dict & "bo伯,bo驳,bo玻,bo剥,bo薄,bo勃,bo菠,bo钵,bo搏,bo脖,bo帛,bo柏,bo舶,bo渤,bo铂,bo箔,bo膊,bo卜,bo礴,bo跛,bo檗,bo亳,bo鹁,bo踣,bo啵,bo蕃,bo簸,bo钹,bo饽,bo擘,bo孛,bo百,bo趵,bu不,bu步,bu补,bu布,bu部,bu捕,bu卜,"
		Pinyin_Dict = Pinyin_Dict & "bu簿,bu哺,bu堡,bu埠,bu怖,bu埔,bu瓿,bu逋,bu晡,bu钸,bu钚,bu醭,bu卟,ca擦,ca礤,ca嚓,cai才,cai菜,cai采,cai材,cai财,cai裁,cai猜,cai踩,cai睬,cai蔡,cai彩,can蚕,can残,can掺,can参,can惨,can惭,can餐,can灿,can骖,can璨,can孱,can黪,can粲,"
		Pinyin_Dict = Pinyin_Dict & "cang藏,cang仓,cang沧,cang舱,cang苍,cang伧,cao草,cao操,cao曹,cao槽,cao糙,cao嘈,cao艚,cao螬,cao漕,cao艹,ce册,ce侧,ce策,ce测,ce厕,ce恻,cen参,cen岑,cen涔,ceng曾,ceng层,ceng蹭,ceng噌,cha查,cha插,cha叉,cha茶,cha差,cha岔,cha搽,cha察,cha茬,cha碴,cha刹,"
		Pinyin_Dict = Pinyin_Dict & "cha诧,cha楂,cha槎,cha镲,cha衩,cha汊,cha馇,cha檫,cha姹,cha杈,cha锸,cha嚓,cha猹,chai柴,chai拆,chai差,chai豺,chai钗,chai瘥,chai虿,chai侪,chai龇,chan产,chan缠,chan掺,chan搀,chan阐,chan颤,chan铲,chan谗,chan蝉,chan单,chan馋,chan觇,chan婵,chan蒇,chan谄,chan冁,chan廛,chan孱,"
		Pinyin_Dict = Pinyin_Dict & "chan蟾,chan羼,chan镡,chan忏,chan潺,chan禅,chan躔,chan澶,chang长,chang唱,chang常,chang场,chang厂,chang尝,chang肠,chang畅,chang昌,chang敞,chang倡,chang偿,chang猖,chang鲳,chang氅,chang菖,chang惝,chang嫦,chang徜,chang鬯,chang阊,chang怅,chang伥,chang昶,chang苌,chang娼,chao朝,chao抄,chao超,chao吵,chao潮,chao巢,"
		Pinyin_Dict = Pinyin_Dict & "chao炒,chao嘲,chao剿,chao绰,chao钞,chao怊,chao耖,chao晁,che车,che撤,che扯,che掣,che彻,che澈,che坼,che砗,che屮,chen趁,chen称,chen辰,chen臣,chen尘,chen晨,chen沉,chen陈,chen衬,chen忱,chen郴,chen榇,chen抻,chen谌,chen碜,chen谶,chen宸,chen龀,chen嗔,chen琛,cheng成,cheng乘,cheng盛,"
		Pinyin_Dict = Pinyin_Dict & "cheng撑,cheng称,cheng城,cheng程,cheng呈,cheng诚,cheng秤,cheng惩,cheng逞,cheng骋,cheng澄,cheng橙,cheng承,cheng塍,cheng柽,cheng埕,cheng铖,cheng噌,cheng铛,cheng酲,cheng裎,cheng枨,cheng蛏,cheng丞,cheng瞠,cheng徵,chi吃,chi尺,chi迟,chi池,chi翅,chi痴,chi赤,chi齿,chi耻,chi持,chi斥,chi侈,chi弛,chi驰,"
		Pinyin_Dict = Pinyin_Dict & "chi炽,chi匙,chi踟,chi坻,chi茌,chi墀,chi饬,chi媸,chi豉,chi褫,chi敕,chi哧,chi瘛,chi蚩,chi啻,chi鸱,chi眵,chi螭,chi篪,chi魑,chi叱,chi彳,chi笞,chi嗤,chi傺,chong冲,chong重,chong虫,chong充,chong宠,chong崇,chong种,chong艟,chong忡,chong舂,chong铳,chong憧,chong茺,chou抽,chou愁,"
		Pinyin_Dict = Pinyin_Dict & "chou臭,chou仇,chou丑,chou稠,chou绸,chou酬,chou筹,chou踌,chou畴,chou瞅,chou惆,chou俦,chou帱,chou瘳,chou雠,chu出,chu处,chu初,chu锄,chu除,chu触,chu橱,chu楚,chu础,chu储,chu畜,chu滁,chu矗,chu搐,chu躇,chu厨,chu雏,chu楮,chu杵,chu刍,chu怵,chu绌,chu亍,chu憷,chu蹰,"
		Pinyin_Dict = Pinyin_Dict & "chu黜,chu蜍,chu樗,chu褚,chuai揣,chuai膪,chuai嘬,chuai搋,chuai踹,chuan穿,chuan船,chuan传,chuan串,chuan川,chuan喘,chuan椽,chuan氚,chuan遄,chuan钏,chuan舡,chuan舛,chuan巛,chuang窗,chuang床,chuang闯,chuang创,chuang疮,chuang幢,chuang怆,chui吹,chui垂,chui炊,chui锤,chui捶,chui槌,chui棰,chui陲,chun春,chun唇,chun纯,"
		Pinyin_Dict = Pinyin_Dict & "chun蠢,chun醇,chun淳,chun椿,chun蝽,chun莼,chun鹑,chuo戳,chuo绰,chuo踔,chuo啜,chuo龊,chuo辍,chuo辶,ci次,ci此,ci词,ci瓷,ci慈,ci雌,ci磁,ci辞,ci刺,ci茨,ci伺,ci疵,ci赐,ci差,ci兹,ci呲,ci鹚,ci祠,ci糍,ci粢,ci茈,cong从,cong丛,cong葱,cong匆,cong聪,"
		Pinyin_Dict = Pinyin_Dict & "cong囱,cong琮,cong枞,cong淙,cong璁,cong骢,cong苁,cou凑,cou楱,cou辏,cou腠,cu粗,cu醋,cu簇,cu促,cu徂,cu猝,cu蔟,cu蹙,cu酢,cu殂,cu蹴,cuan窜,cuan蹿,cuan篡,cuan攒,cuan汆,cuan爨,cuan镩,cuan撺,cui催,cui脆,cui摧,cui翠,cui崔,cui淬,cui瘁,cui粹,cui璀,cui啐,"
		Pinyin_Dict = Pinyin_Dict & "cui悴,cui萃,cui毳,cui榱,cui隹,cun村,cun寸,cun存,cun忖,cun皴,cuo错,cuo撮,cuo搓,cuo挫,cuo措,cuo磋,cuo嵯,cuo厝,cuo鹾,cuo脞,cuo痤,cuo蹉,cuo瘥,cuo锉,cuo矬,cuo躜,da大,da答,da达,da打,da搭,da瘩,da笪,da耷,da哒,da褡,da疸,da怛,da靼,da妲,"
		Pinyin_Dict = Pinyin_Dict & "da沓,da嗒,da鞑,dai带,dai代,dai呆,dai戴,dai待,dai袋,dai逮,dai歹,dai贷,dai怠,dai傣,dai大,dai殆,dai呔,dai玳,dai迨,dai岱,dai甙,dai黛,dai骀,dai绐,dai埭,dan但,dan单,dan蛋,dan担,dan弹,dan掸,dan胆,dan淡,dan丹,dan耽,dan旦,dan氮,dan诞,dan郸,dan惮,"
		Pinyin_Dict = Pinyin_Dict & "dan石,dan疸,dan澹,dan瘅,dan萏,dan殚,dan眈,dan聃,dan箪,dan赕,dan儋,dan啖,dan赡,dang当,dang党,dang挡,dang档,dang荡,dang谠,dang铛,dang宕,dang菪,dang凼,dang裆,dang砀,dao到,dao道,dao倒,dao刀,dao岛,dao盗,dao稻,dao捣,dao悼,dao导,dao蹈,dao祷,dao帱,dao纛,dao忉,"
		Pinyin_Dict = Pinyin_Dict & "dao焘,dao氘,dao叨,dao刂,de的,de地,de得,de德,de锝,dei得,deng等,deng灯,deng邓,deng登,deng澄,deng瞪,deng凳,deng蹬,deng磴,deng镫,deng噔,deng嶝,deng戥,deng簦,di地,di第,di底,di低,di敌,di抵,di滴,di帝,di递,di嫡,di弟,di缔,di堤,di的,di涤,di提,"
		Pinyin_Dict = Pinyin_Dict & "di笛,di迪,di狄,di翟,di蒂,di觌,di邸,di谛,di诋,di嘀,di柢,di骶,di羝,di氐,di棣,di睇,di娣,di荻,di碲,di镝,di坻,di籴,di砥,dia嗲,dian点,dian电,dian店,dian殿,dian淀,dian掂,dian颠,dian垫,dian碘,dian惦,dian奠,dian典,dian佃,dian靛,dian滇,dian甸,"
		Pinyin_Dict = Pinyin_Dict & "dian踮,dian钿,dian坫,dian阽,dian癫,dian簟,dian玷,dian巅,dian癜,diao掉,diao钓,diao叼,diao吊,diao雕,diao调,diao刁,diao碉,diao凋,diao铞,diao铫,diao鲷,diao貂,die爹,die跌,die叠,die碟,die蝶,die迭,die谍,die牒,die堞,die瓞,die揲,die蹀,die耋,die鲽,die垤,die喋,ding顶,ding定,"
		Pinyin_Dict = Pinyin_Dict & "ding盯,ding订,ding叮,ding丁,ding钉,ding鼎,ding锭,ding町,ding玎,ding铤,ding腚,ding碇,ding疔,ding仃,ding耵,ding酊,ding啶,diu丢,diu铥,dong动,dong东,dong懂,dong洞,dong冻,dong冬,dong董,dong栋,dong侗,dong恫,dong峒,dong鸫,dong胨,dong胴,dong硐,dong氡,dong岽,dong咚,dou都,dou斗,dou豆,"
		Pinyin_Dict = Pinyin_Dict & "dou逗,dou陡,dou抖,dou痘,dou兜,dou蚪,dou窦,dou篼,dou蔸,du读,du度,du毒,du渡,du堵,du独,du肚,du镀,du赌,du睹,du杜,du督,du都,du犊,du妒,du蠹,du笃,du嘟,du渎,du椟,du牍,du黩,du髑,du芏,duan段,duan短,duan断,duan端,duan锻,duan缎,duan椴,"
		Pinyin_Dict = Pinyin_Dict & "duan煅,duan簖,dui对,dui队,dui堆,dui兑,dui碓,dui怼,dui憝,dun吨,dun顿,dun蹲,dun墩,dun敦,dun钝,dun盾,dun囤,dun遁,dun趸,dun沌,dun盹,dun镦,dun礅,dun炖,dun砘,duo多,duo朵,duo夺,duo舵,duo剁,duo垛,duo跺,duo惰,duo堕,duo掇,duo哆,duo驮,duo度,duo躲,duo踱,"
		Pinyin_Dict = Pinyin_Dict & "duo沲,duo咄,duo铎,duo裰,duo哚,duo缍,e饿,e哦,e额,e鹅,e蛾,e扼,e俄,e讹,e阿,e遏,e峨,e娥,e恶,e厄,e鄂,e锇,e谔,e垩,e锷,e萼,e苊,e轭,e婀,e莪,e鳄,e颚,e腭,e愕,e呃,e噩,e鹗,e屙,ei诶,en恩,"
		Pinyin_Dict = Pinyin_Dict & "en摁,en蒽,er而,er二,er耳,er儿,er饵,er尔,er贰,er洱,er珥,er鲕,er鸸,er迩,er铒,fa发,fa法,fa罚,fa伐,fa乏,fa筏,fa阀,fa珐,fa垡,fa砝,fan反,fan饭,fan翻,fan番,fan犯,fan凡,fan帆,fan返,fan泛,fan繁,fan烦,fan贩,fan范,fan樊,fan藩,"
		Pinyin_Dict = Pinyin_Dict & "fan矾,fan钒,fan燔,fan蘩,fan畈,fan蕃,fan蹯,fan梵,fan幡,fang放,fang房,fang防,fang纺,fang芳,fang方,fang访,fang仿,fang坊,fang妨,fang肪,fang钫,fang邡,fang枋,fang舫,fang鲂,fang匚,fei非,fei飞,fei肥,fei费,fei肺,fei废,fei匪,fei吠,fei沸,fei菲,fei诽,fei啡,fei篚,fei蜚,"
		Pinyin_Dict = Pinyin_Dict & "fei腓,fei扉,fei妃,fei斐,fei狒,fei芾,fei悱,fei镄,fei霏,fei翡,fei榧,fei淝,fei鲱,fei绯,fei痱,fei砩,fen分,fen份,fen芬,fen粉,fen坟,fen奋,fen愤,fen纷,fen忿,fen粪,fen酚,fen焚,fen吩,fen氛,fen汾,fen棼,fen瀵,fen鲼,fen玢,fen偾,fen鼢,fen贲,feng风,feng封,"
		Pinyin_Dict = Pinyin_Dict & "feng逢,feng缝,feng蜂,feng丰,feng枫,feng疯,feng冯,feng奉,feng讽,feng凤,feng峰,feng锋,feng烽,feng砜,feng俸,feng酆,feng葑,feng沣,feng唪,fo佛,fou否,fou缶,fu副,fu幅,fu扶,fu浮,fu富,fu福,fu负,fu伏,fu付,fu复,fu服,fu附,fu俯,fu斧,fu赴,fu缚,fu拂,fu夫,"
		Pinyin_Dict = Pinyin_Dict & "fu父,fu符,fu孵,fu敷,fu赋,fu辅,fu府,fu腐,fu腹,fu妇,fu抚,fu覆,fu辐,fu肤,fu氟,fu佛,fu俘,fu傅,fu讣,fu弗,fu涪,fu袱,fu甫,fu釜,fu脯,fu腑,fu阜,fu咐,fu黼,fu砩,fu苻,fu趺,fu跗,fu蚨,fu芾,fu鲋,fu幞,fu茯,fu滏,fu蜉,"
		Pinyin_Dict = Pinyin_Dict & "fu拊,fu菔,fu蝠,fu鳆,fu蝮,fu绂,fu绋,fu赙,fu罘,fu稃,fu匐,fu麸,fu凫,fu桴,fu莩,fu孚,fu馥,fu驸,fu怫,fu祓,fu呋,fu郛,fu芙,fu艴,fu黻,fu哺,fu阝,ga噶,ga夹,ga嘎,ga咖,ga钆,ga伽,ga旮,ga尬,ga尕,ga尜,gai该,gai改,gai盖,"
		Pinyin_Dict = Pinyin_Dict & "gai概,gai钙,gai芥,gai溉,gai戤,gai垓,gai丐,gai陔,gai赅,gai胲,gan赶,gan干,gan感,gan敢,gan竿,gan甘,gan肝,gan柑,gan杆,gan赣,gan秆,gan旰,gan酐,gan矸,gan疳,gan泔,gan苷,gan擀,gan绀,gan橄,gan澉,gan淦,gan尴,gan坩,gang刚,gang钢,gang纲,gang港,gang缸,gang岗,"
		Pinyin_Dict = Pinyin_Dict & "gang杠,gang冈,gang肛,gang扛,gang筻,gang罡,gang戆,gao高,gao搞,gao告,gao稿,gao膏,gao篙,gao羔,gao糕,gao镐,gao皋,gao郜,gao诰,gao杲,gao缟,gao睾,gao槔,gao锆,gao槁,gao藁,ge个,ge各,ge歌,ge割,ge哥,ge搁,ge格,ge阁,ge隔,ge革,ge咯,ge胳,ge葛,ge蛤,"
		Pinyin_Dict = Pinyin_Dict & "ge戈,ge鸽,ge疙,ge盖,ge合,ge铬,ge骼,ge袼,ge塥,ge虼,ge圪,ge镉,ge仡,ge舸,ge鬲,ge嗝,ge膈,ge搿,ge纥,ge哿,ge铪,gei给,gen跟,gen根,gen哏,gen茛,gen亘,gen艮,geng更,geng耕,geng颈,geng梗,geng耿,geng庚,geng羹,geng埂,geng赓,geng鲠,geng哽,geng绠,"
		Pinyin_Dict = Pinyin_Dict & "gong工,gong公,gong功,gong共,gong弓,gong攻,gong宫,gong供,gong恭,gong拱,gong贡,gong躬,gong巩,gong汞,gong龚,gong肱,gong觥,gong珙,gong蚣,gong廾,gou够,gou沟,gou狗,gou钩,gou勾,gou购,gou构,gou苟,gou垢,gou岣,gou彀,gou枸,gou鞲,gou觏,gou缑,gou笱,gou诟,gou遘,gou媾,gou篝,"
		Pinyin_Dict = Pinyin_Dict & "gou佝,gu古,gu股,gu鼓,gu谷,gu故,gu孤,gu箍,gu姑,gu顾,gu固,gu雇,gu估,gu咕,gu骨,gu辜,gu沽,gu蛊,gu贾,gu菇,gu梏,gu鸪,gu汩,gu轱,gu崮,gu菰,gu鹄,gu鹘,gu钴,gu臌,gu酤,gu鲴,gu诂,gu牯,gu瞽,gu毂,gu锢,gu牿,gu痼,gu觚,"
		Pinyin_Dict = Pinyin_Dict & "gu蛄,gu罟,gu嘏,gua挂,gua刮,gua瓜,gua寡,gua剐,gua褂,gua卦,gua呱,gua胍,gua鸹,gua栝,gua诖,guai怪,guai拐,guai乖,guai掴,guan关,guan管,guan官,guan观,guan馆,guan惯,guan罐,guan灌,guan冠,guan贯,guan棺,guan纶,guan盥,guan莞,guan掼,guan涫,guan鳏,guan鹳,guan倌,guang光,guang广,"
		Pinyin_Dict = Pinyin_Dict & "guang逛,guang桄,guang犷,guang咣,guang胱,gui归,gui贵,gui鬼,gui跪,gui轨,gui规,gui硅,gui桂,gui柜,gui龟,gui诡,gui闺,gui瑰,gui圭,gui刽,gui癸,gui炔,gui庋,gui宄,gui桧,gui刿,gui鳜,gui鲑,gui皈,gui匦,gui妫,gui晷,gui簋,gui炅,gun滚,gun棍,gun辊,gun鲧,gun衮,gun磙,"
		Pinyin_Dict = Pinyin_Dict & "gun绲,guo过,guo国,guo果,guo裹,guo锅,guo郭,guo涡,guo埚,guo椁,guo聒,guo馘,guo猓,guo崞,guo掴,guo帼,guo呙,guo虢,guo蜾,guo蝈,guo锞,ha哈,ha蛤,ha铪,hai还,hai海,hai害,hai咳,hai氦,hai孩,hai骇,hai咳,hai骸,hai亥,hai嗨,hai醢,hai胲,han喊,han含,han汗,"
		Pinyin_Dict = Pinyin_Dict & "han寒,han汉,han旱,han酣,han韩,han焊,han涵,han函,han憨,han翰,han罕,han撼,han捍,han憾,han悍,han邯,han邗,han菡,han撖,han瀚,han顸,han蚶,han焓,han颔,han晗,han鼾,hang行,hang巷,hang航,hang夯,hang杭,hang吭,hang颃,hang沆,hang绗,hao好,hao号,hao浩,hao嚎,hao壕,"
		Pinyin_Dict = Pinyin_Dict & "hao郝,hao毫,hao豪,hao耗,hao貉,hao镐,hao昊,hao颢,hao灏,hao嚆,hao蚝,hao嗥,hao皓,hao蒿,hao濠,hao薅,he和,he喝,he合,he河,he禾,he核,he何,he呵,he荷,he贺,he赫,he褐,he盒,he鹤,he菏,he貉,he阂,he涸,he吓,he嗬,he劾,he盍,he翮,he阖,"
		Pinyin_Dict = Pinyin_Dict & "he颌,he壑,he诃,he纥,he曷,he蚵,hei黑,hei嘿,hen很,hen狠,hen恨,hen痕,heng横,heng恒,heng哼,heng衡,heng亨,heng桁,heng珩,heng蘅,hong红,hong轰,hong哄,hong虹,hong洪,hong宏,hong烘,hong鸿,hong弘,hong讧,hong訇,hong蕻,hong闳,hong薨,hong黉,hong荭,hong泓,hou后,hou厚,hou吼,"
		Pinyin_Dict = Pinyin_Dict & "hou喉,hou侯,hou候,hou猴,hou鲎,hou篌,hou堠,hou後,hou逅,hou糇,hou骺,hou瘊,hu湖,hu户,hu呼,hu虎,hu壶,hu互,hu胡,hu护,hu糊,hu弧,hu忽,hu狐,hu蝴,hu葫,hu沪,hu乎,hu核,hu瑚,hu唬,hu鹕,hu冱,hu怙,hu鹱,hu笏,hu戽,hu扈,hu鹘,hu浒,"
		Pinyin_Dict = Pinyin_Dict & "hu祜,hu醐,hu琥,hu囫,hu烀,hu轷,hu瓠,hu煳,hu斛,hu鹄,hu猢,hu惚,hu岵,hu滹,hu觳,hu唿,hu槲,hu虍,hua话,hua花,hua化,hua画,hua华,hua划,hua滑,hua哗,hua猾,hua铧,hua桦,hua骅,hua砉,huai坏,huai怀,huai淮,huai槐,huai徊,huai踝,huan换,huan还,huan唤,"
		Pinyin_Dict = Pinyin_Dict & "huan环,huan患,huan缓,huan欢,huan幻,huan宦,huan涣,huan焕,huan豢,huan桓,huan痪,huan漶,huan獾,huan擐,huan逭,huan鲩,huan郇,huan鬟,huan寰,huan奂,huan锾,huan圜,huan洹,huan萑,huan缳,huan浣,huan垸,huang黄,huang慌,huang晃,huang荒,huang簧,huang凰,huang皇,huang谎,huang惶,huang蝗,huang磺,huang恍,huang煌,"
		Pinyin_Dict = Pinyin_Dict & "huang幌,huang隍,huang肓,huang潢,huang篁,huang徨,huang鳇,huang遑,huang癀,huang湟,huang蟥,huang璜,hui回,hui会,hui灰,hui绘,hui挥,hui汇,hui辉,hui毁,hui悔,hui惠,hui晦,hui徽,hui恢,hui秽,hui慧,hui贿,hui蛔,hui讳,hui卉,hui烩,hui诲,hui彗,hui浍,hui蕙,hui喙,hui恚,hui哕,hui晖,"
		Pinyin_Dict = Pinyin_Dict & "hui隳,hui麾,hui诙,hui蟪,hui茴,hui洄,hui咴,hui虺,hui荟,hui缋,hui桧,hun混,hun昏,hun荤,hun浑,hun婚,hun魂,hun阍,hun珲,hun馄,hun溷,hun诨,huo或,huo活,huo火,huo伙,huo货,huo和,huo获,huo祸,huo豁,huo霍,huo惑,huo嚯,huo镬,huo耠,huo劐,huo藿,huo攉,huo锪,"
		Pinyin_Dict = Pinyin_Dict & "huo蠖,huo钬,huo夥,ji几,ji及,ji急,ji既,ji即,ji机,ji鸡,ji积,ji记,ji级,ji极,ji计,ji挤,ji己,ji季,ji寄,ji纪,ji系,ji基,ji激,ji吉,ji脊,ji际,ji汲,ji肌,ji嫉,ji姬,ji绩,ji缉,ji饥,ji迹,ji棘,ji蓟,ji技,ji冀,ji辑,ji伎,"
		Pinyin_Dict = Pinyin_Dict & "ji祭,ji剂,ji悸,ji济,ji籍,ji寂,ji奇,ji忌,ji妓,ji继,ji集,ji给,ji击,ji圾,ji箕,ji讥,ji畸,ji稽,ji疾,ji墼,ji洎,ji鲚,ji屐,ji齑,ji戟,ji鲫,ji嵇,ji矶,ji稷,ji戢,ji虮,ji笈,ji暨,ji笄,ji剞,ji叽,ji蒺,ji跻,ji嵴,ji掎,"
		Pinyin_Dict = Pinyin_Dict & "ji跽,ji霁,ji唧,ji畿,ji荠,ji瘠,ji玑,ji羁,ji丌,ji偈,ji芨,ji佶,ji赍,ji楫,ji髻,ji咭,ji蕺,ji觊,ji麂,ji骥,ji殛,ji岌,ji亟,ji犄,ji乩,ji芰,ji哜,ji彐,ji萁,ji藉,jia家,jia加,jia假,jia价,jia架,jia甲,jia佳,jia夹,jia嘉,jia驾,"
		Pinyin_Dict = Pinyin_Dict & "jia嫁,jia枷,jia荚,jia颊,jia钾,jia稼,jia茄,jia贾,jia铗,jia葭,jia迦,jia戛,jia浃,jia镓,jia痂,jia恝,jia岬,jia跏,jia嘏,jia伽,jia胛,jia笳,jia珈,jia瘕,jia郏,jia袈,jia蛱,jia袷,jia铪,jian见,jian件,jian减,jian尖,jian间,jian键,jian贱,jian肩,jian兼,jian建,jian检,"
		Pinyin_Dict = Pinyin_Dict & "jian箭,jian煎,jian简,jian剪,jian歼,jian监,jian坚,jian奸,jian健,jian艰,jian荐,jian剑,jian渐,jian溅,jian涧,jian鉴,jian践,jian捡,jian柬,jian笺,jian俭,jian碱,jian硷,jian拣,jian舰,jian槛,jian缄,jian茧,jian饯,jian翦,jian鞯,jian戋,jian谏,jian牮,jian枧,jian腱,jian趼,jian缣,jian搛,jian戬,"
		Pinyin_Dict = Pinyin_Dict & "jian毽,jian菅,jian鲣,jian笕,jian谫,jian楗,jian囝,jian蹇,jian裥,jian踺,jian睑,jian謇,jian鹣,jian蒹,jian僭,jian锏,jian湔,jian犍,jian谮,jiang将,jiang讲,jiang江,jiang奖,jiang降,jiang浆,jiang僵,jiang姜,jiang酱,jiang蒋,jiang疆,jiang匠,jiang强,jiang桨,jiang虹,jiang豇,jiang礓,jiang缰,jiang犟,jiang耩,jiang绛,"
		Pinyin_Dict = Pinyin_Dict & "jiang茳,jiang糨,jiang洚,jiao叫,jiao脚,jiao交,jiao角,jiao教,jiao较,jiao缴,jiao觉,jiao焦,jiao胶,jiao娇,jiao绞,jiao校,jiao搅,jiao骄,jiao狡,jiao浇,jiao矫,jiao郊,jiao嚼,jiao蕉,jiao轿,jiao窖,jiao椒,jiao礁,jiao饺,jiao铰,jiao酵,jiao侥,jiao剿,jiao徼,jiao艽,jiao僬,jiao蛟,jiao敫,jiao峤,jiao跤,"
		Pinyin_Dict = Pinyin_Dict & "jiao姣,jiao皎,jiao茭,jiao鹪,jiao噍,jiao醮,jiao佼,jiao鲛,jiao挢,jie接,jie节,jie街,jie借,jie皆,jie截,jie解,jie界,jie结,jie届,jie姐,jie揭,jie戒,jie介,jie阶,jie劫,jie芥,jie竭,jie洁,jie疥,jie藉,jie秸,jie桔,jie杰,jie捷,jie诫,jie睫,jie偈,jie桀,jie喈,jie拮,"
		Pinyin_Dict = Pinyin_Dict & "jie骱,jie羯,jie蚧,jie嗟,jie颉,jie鲒,jie婕,jie碣,jie讦,jie孑,jie疖,jie诘,jie卩,jie锴,jin进,jin近,jin今,jin仅,jin紧,jin金,jin斤,jin尽,jin劲,jin禁,jin浸,jin锦,jin晋,jin筋,jin津,jin谨,jin巾,jin襟,jin烬,jin靳,jin廑,jin瑾,jin馑,jin槿,jin衿,jin堇,"
		Pinyin_Dict = Pinyin_Dict & "jin荩,jin矜,jin噤,jin缙,jin卺,jin妗,jin赆,jin觐,jin钅,jing竟,jing静,jing井,jing惊,jing经,jing镜,jing京,jing净,jing敬,jing精,jing景,jing警,jing竞,jing境,jing径,jing荆,jing晶,jing鲸,jing粳,jing颈,jing兢,jing茎,jing睛,jing劲,jing痉,jing靖,jing肼,jing獍,jing阱,jing腈,jing弪,"
		Pinyin_Dict = Pinyin_Dict & "jing刭,jing憬,jing婧,jing胫,jing菁,jing儆,jing旌,jing迳,jing靓,jing泾,jing陉,jiong窘,jiong炯,jiong扃,jiong迥,jiong冂,jiu就,jiu九,jiu酒,jiu旧,jiu久,jiu揪,jiu救,jiu纠,jiu舅,jiu究,jiu韭,jiu厩,jiu臼,jiu玖,jiu灸,jiu咎,jiu疚,jiu赳,jiu鹫,jiu僦,jiu柩,jiu桕,jiu鬏,jiu鸠,"
		Pinyin_Dict = Pinyin_Dict & "jiu阄,jiu啾,ju句,ju举,ju巨,ju局,ju具,ju距,ju锯,ju剧,ju居,ju聚,ju拘,ju菊,ju矩,ju沮,ju拒,ju惧,ju鞠,ju狙,ju驹,ju据,ju柜,ju俱,ju车,ju咀,ju疽,ju踞,ju炬,ju倨,ju醵,ju裾,ju屦,ju犋,ju苴,ju窭,ju飓,ju锔,ju椐,ju苣,"
		Pinyin_Dict = Pinyin_Dict & "ju琚,ju掬,ju榘,ju龃,ju趄,ju莒,ju雎,ju遽,ju橘,ju踽,ju榉,ju鞫,ju钜,ju讵,ju枸,ju瞿,ju蘧,juan卷,juan圈,juan倦,juan鹃,juan捐,juan娟,juan眷,juan绢,juan鄄,juan锩,juan蠲,juan镌,juan狷,juan桊,juan涓,juan隽,jue决,jue绝,jue觉,jue角,jue爵,jue掘,jue诀,"
		Pinyin_Dict = Pinyin_Dict & "jue撅,jue倔,jue抉,jue攫,jue嚼,jue脚,jue桷,jue橛,jue觖,jue劂,jue爝,jue矍,jue镢,jue獗,jue珏,jue崛,jue蕨,jue噘,jue谲,jue蹶,jue孓,jue厥,jue阙,jun军,jun君,jun均,jun菌,jun俊,jun峻,jun龟,jun竣,jun骏,jun钧,jun浚,jun郡,jun筠,jun麇,jun皲,jun捃,ka卡,"
		Pinyin_Dict = Pinyin_Dict & "ka喀,ka咯,ka咖,ka胩,ka咔,ka佧,kai开,kai揩,kai凯,kai慨,kai楷,kai垲,kai剀,kai锎,kai铠,kai锴,kai忾,kai恺,kai蒈,kan看,kan砍,kan堪,kan刊,kan坎,kan槛,kan勘,kan龛,kan戡,kan侃,kan瞰,kan莰,kan阚,kan凵,kang抗,kang炕,kang扛,kang糠,kang康,kang慷,kang亢,"
		Pinyin_Dict = Pinyin_Dict & "kang钪,kang闶,kang伉,kao靠,kao考,kao烤,kao拷,kao栲,kao犒,kao尻,kao铐,ke可,ke克,ke棵,ke科,ke颗,ke刻,ke课,ke客,ke壳,ke渴,ke苛,ke柯,ke磕,ke咳,ke坷,ke呵,ke恪,ke岢,ke蝌,ke缂,ke蚵,ke轲,ke窠,ke钶,ke氪,ke颏,ke瞌,ke锞,ke稞,"
		Pinyin_Dict = Pinyin_Dict & "ke珂,ke髁,ke疴,ke嗑,ke溘,ke骒,ke铪,ken肯,ken啃,ken恳,ken垦,ken裉,keng坑,keng吭,keng铿,keng胫,keng铒,kong空,kong孔,kong控,kong恐,kong倥,kong崆,kong箜,kou口,kou扣,kou抠,kou寇,kou蔻,kou芤,kou眍,kou筘,kou叩,ku哭,ku库,ku苦,ku枯,ku裤,ku窟,ku酷,"
		Pinyin_Dict = Pinyin_Dict & "ku刳,ku骷,ku喾,ku堀,ku绔,kua跨,kua垮,kua挎,kua夸,kua胯,kua侉,kua锞,kuai快,kuai块,kuai筷,kuai会,kuai侩,kuai哙,kuai蒯,kuai郐,kuai狯,kuai脍,kuan宽,kuan款,kuan髋,kuang矿,kuang筐,kuang狂,kuang框,kuang况,kuang旷,kuang匡,kuang眶,kuang诳,kuang邝,kuang纩,kuang夼,kuang诓,kuang圹,kuang贶,"
		Pinyin_Dict = Pinyin_Dict & "kuang哐,kui亏,kui愧,kui奎,kui窥,kui溃,kui葵,kui魁,kui馈,kui盔,kui傀,kui岿,kui匮,kui愦,kui揆,kui睽,kui跬,kui聩,kui篑,kui喹,kui逵,kui暌,kui蒉,kui悝,kui喟,kui馗,kui蝰,kui隗,kui夔,kun捆,kun困,kun昆,kun坤,kun鲲,kun锟,kun髡,kun琨,kun醌,kun阃,kun悃,"
		Pinyin_Dict = Pinyin_Dict & "kun顽,kuo阔,kuo扩,kuo括,kuo廓,kuo蛞,la拉,la啦,la辣,la蜡,la腊,la喇,la垃,la落,la瘌,la邋,la砬,la剌,la旯,lai来,lai赖,lai莱,lai濑,lai赉,lai崃,lai涞,lai铼,lai籁,lai徕,lai癞,lai睐,lan蓝,lan兰,lan烂,lan拦,lan篮,lan懒,lan栏,lan揽,lan缆,"
		Pinyin_Dict = Pinyin_Dict & "lan滥,lan阑,lan谰,lan婪,lan澜,lan览,lan榄,lan岚,lan褴,lan镧,lan斓,lan罱,lan漤,lang浪,lang狼,lang廊,lang郎,lang朗,lang榔,lang琅,lang稂,lang螂,lang莨,lang啷,lang锒,lang阆,lang蒗,lao老,lao捞,lao牢,lao劳,lao烙,lao涝,lao落,lao姥,lao酪,lao络,lao佬,lao耢,lao铹,"
		Pinyin_Dict = Pinyin_Dict & "lao醪,lao铑,lao唠,lao栳,lao崂,lao痨,le了,le乐,le勒,le鳓,le仂,le叻,le泐,lei类,lei累,lei泪,lei雷,lei垒,lei勒,lei擂,lei蕾,lei肋,lei镭,lei儡,lei磊,lei缧,lei诔,lei耒,lei酹,lei羸,lei嫘,lei檑,lei嘞,leng冷,leng棱,leng楞,leng愣,leng塄,li里,li离,"
		Pinyin_Dict = Pinyin_Dict & "li力,li立,li李,li例,li哩,li理,li利,li梨,li厘,li礼,li历,li丽,li吏,li砾,li漓,li莉,li傈,li荔,li俐,li痢,li狸,li粒,li沥,li隶,li栗,li璃,li鲤,li厉,li励,li犁,li黎,li篱,li郦,li鹂,li笠,li坜,li苈,li鳢,li缡,li跞,"
		Pinyin_Dict = Pinyin_Dict & "li蜊,li锂,li澧,li粝,li蓠,li枥,li蠡,li鬲,li呖,li砺,li嫠,li篥,li疠,li疬,li猁,li藜,li溧,li鲡,li戾,li栎,li唳,li醴,li轹,li詈,li骊,li罹,li逦,li俪,li喱,li雳,li黧,li莅,li俚,li蛎,li娌,li砬,lia俩,lian连,lian联,lian练,"
		Pinyin_Dict = Pinyin_Dict & "lian莲,lian恋,lian脸,lian炼,lian链,lian敛,lian怜,lian廉,lian帘,lian镰,lian涟,lian蠊,lian琏,lian殓,lian蔹,lian鲢,lian奁,lian潋,lian臁,lian裢,lian濂,lian裣,lian楝,liang两,liang亮,liang辆,liang凉,liang粮,liang梁,liang量,liang良,liang晾,liang谅,liang俩,liang粱,liang墚,liang踉,liang椋,liang魉,liang莨,"
		Pinyin_Dict = Pinyin_Dict & "liao了,liao料,liao撩,liao聊,liao撂,liao疗,liao廖,liao燎,liao辽,liao僚,liao寥,liao镣,liao潦,liao钌,liao蓼,liao尥,liao寮,liao缭,liao獠,liao鹩,liao嘹,lie列,lie裂,lie猎,lie劣,lie烈,lie咧,lie埒,lie捩,lie鬣,lie趔,lie躐,lie冽,lie洌,lin林,lin临,lin淋,lin邻,lin磷,lin鳞,"
		Pinyin_Dict = Pinyin_Dict & "lin赁,lin吝,lin拎,lin琳,lin霖,lin凛,lin遴,lin嶙,lin蔺,lin粼,lin麟,lin躏,lin辚,lin廪,lin懔,lin瞵,lin檩,lin膦,lin啉,ling另,ling令,ling领,ling零,ling铃,ling玲,ling灵,ling岭,ling龄,ling凌,ling陵,ling菱,ling伶,ling羚,ling棱,ling翎,ling蛉,ling苓,ling绫,ling瓴,ling酃,"
		Pinyin_Dict = Pinyin_Dict & "ling呤,ling泠,ling棂,ling柃,ling鲮,ling聆,ling囹,liu六,liu流,liu留,liu刘,liu柳,liu溜,liu硫,liu瘤,liu榴,liu琉,liu馏,liu碌,liu陆,liu绺,liu锍,liu鎏,liu镏,liu浏,liu骝,liu旒,liu鹨,liu熘,liu遛,lo咯,long龙,long拢,long笼,long聋,long隆,long垄,long弄,long咙,long窿,"
		Pinyin_Dict = Pinyin_Dict & "long陇,long垅,long胧,long珑,long茏,long泷,long栊,long癃,long砻,lou楼,lou搂,lou漏,lou陋,lou露,lou娄,lou篓,lou偻,lou蝼,lou镂,lou蒌,lou耧,lou髅,lou喽,lou瘘,lou嵝,lu路,lu露,lu录,lu鹿,lu陆,lu炉,lu卢,lu鲁,lu卤,lu芦,lu颅,lu庐,lu碌,lu掳,lu绿,"
		Pinyin_Dict = Pinyin_Dict & "lu虏,lu赂,lu戮,lu潞,lu禄,lu麓,lu六,lu鲈,lu栌,lu渌,lu逯,lu泸,lu轳,lu氇,lu簏,lu橹,lu辂,lu垆,lu胪,lu噜,lu镥,lu辘,lu漉,lu撸,lu璐,lu鸬,lu鹭,lu舻,luan乱,luan卵,luan滦,luan峦,luan孪,luan挛,luan栾,luan銮,luan脔,luan娈,luan鸾,lue略,"
		Pinyin_Dict = Pinyin_Dict & "lue掠,lue锊,lun论,lun轮,lun抡,lun伦,lun沦,lun仑,lun纶,lun囵,luo落,luo罗,luo锣,luo裸,luo骡,luo烙,luo箩,luo螺,luo萝,luo洛,luo骆,luo逻,luo络,luo咯,luo荦,luo漯,luo蠃,luo雒,luo倮,luo硌,luo椤,luo捋,luo脶,luo瘰,luo摞,luo泺,luo珞,luo镙,luo猡,luo铬,"
		Pinyin_Dict = Pinyin_Dict & "lv绿,lv率,lv铝,lv驴,lv旅,lv屡,lv滤,lv吕,lv律,lv氯,lv缕,lv侣,lv虑,lv履,lv偻,lv膂,lv榈,lv闾,lv捋,lv褛,lv稆,lve略,lve掠,lve锊,m呒,ma吗,ma妈,ma马,ma嘛,ma麻,ma骂,ma抹,ma码,ma玛,ma蚂,ma摩,ma唛,ma蟆,ma犸,ma嬷,"
		Pinyin_Dict = Pinyin_Dict & "ma杩,mai买,mai卖,mai迈,mai埋,mai麦,mai脉,mai劢,mai霾,mai荬,man满,man慢,man瞒,man漫,man蛮,man蔓,man曼,man馒,man埋,man谩,man幔,man鳗,man墁,man螨,man镘,man颟,man鞔,man缦,man熳,mang忙,mang芒,mang盲,mang莽,mang茫,mang氓,mang硭,mang邙,mang蟒,mang漭,mao毛,"
		Pinyin_Dict = Pinyin_Dict & "mao冒,mao帽,mao猫,mao矛,mao卯,mao貌,mao茂,mao贸,mao铆,mao锚,mao茅,mao耄,mao茆,mao瑁,mao蝥,mao髦,mao懋,mao昴,mao牦,mao瞀,mao峁,mao袤,mao蟊,mao旄,mao泖,me么,mei没,mei每,mei煤,mei镁,mei美,mei酶,mei妹,mei枚,mei霉,mei玫,mei眉,mei梅,mei寐,mei昧,"
		Pinyin_Dict = Pinyin_Dict & "mei媒,mei媚,mei嵋,mei猸,mei袂,mei湄,mei浼,mei鹛,mei莓,mei魅,mei镅,mei楣,men门,men们,men闷,men懑,men扪,men钔,men焖,meng猛,meng梦,meng蒙,meng锰,meng孟,meng盟,meng檬,meng萌,meng礞,meng蜢,meng勐,meng懵,meng甍,meng蠓,meng虻,meng朦,meng艋,meng艨,meng瞢,mi米,mi密,"
		Pinyin_Dict = Pinyin_Dict & "mi迷,mi眯,mi蜜,mi谜,mi觅,mi秘,mi弥,mi幂,mi靡,mi糜,mi泌,mi醚,mi蘼,mi縻,mi咪,mi汨,mi麋,mi祢,mi猕,mi弭,mi谧,mi芈,mi脒,mi宓,mi敉,mi嘧,mi糸,mi冖,mian面,mian棉,mian免,mian绵,mian眠,mian缅,mian勉,mian冕,mian娩,mian腼,mian湎,mian眄,"
		Pinyin_Dict = Pinyin_Dict & "mian沔,mian渑,mian宀,miao秒,miao苗,miao庙,miao妙,miao描,miao瞄,miao藐,miao渺,miao眇,miao缪,miao缈,miao淼,miao喵,miao杪,miao鹋,miao邈,mie灭,mie蔑,mie咩,mie篾,mie蠛,mie乜,min民,min抿,min敏,min闽,min皿,min悯,min珉,min愍,min缗,min闵,min玟,min苠,min泯,min黾,min鳘,"
		Pinyin_Dict = Pinyin_Dict & "min岷,ming名,ming明,ming命,ming鸣,ming铭,ming螟,ming冥,ming瞑,ming暝,ming茗,ming溟,ming酩,miu谬,miu缪,mo摸,mo磨,mo抹,mo末,mo膜,mo墨,mo没,mo莫,mo默,mo魔,mo模,mo摩,mo摹,mo漠,mo陌,mo蘑,mo脉,mo沫,mo万,mo寞,mo秣,mo瘼,mo殁,mo镆,mo嫫,"
		Pinyin_Dict = Pinyin_Dict & "mo谟,mo蓦,mo貊,mo貘,mo麽,mo茉,mo馍,mo耱,mou某,mou谋,mou牟,mou眸,mou蛑,mou鍪,mou侔,mou缪,mou哞,mu木,mu母,mu亩,mu幕,mu目,mu墓,mu牧,mu牟,mu模,mu穆,mu暮,mu牡,mu拇,mu募,mu慕,mu睦,mu姆,mu钼,mu毪,mu坶,mu沐,mu仫,mu苜,"
		Pinyin_Dict = Pinyin_Dict & "na那,na拿,na哪,na纳,na钠,na娜,na呐,na衲,na捺,na镎,na肭,nai乃,nai耐,nai奶,nai奈,nai氖,nai萘,nai艿,nai柰,nai鼐,nai佴,nan难,nan南,nan男,nan赧,nan囡,nan蝻,nan楠,nan喃,nan腩,nang囊,nang馕,nang曩,nang囔,nang攮,nao闹,nao脑,nao恼,nao挠,nao淖,"
		Pinyin_Dict = Pinyin_Dict & "nao孬,nao铙,nao瑙,nao垴,nao呶,nao蛲,nao猱,nao硇,ne呢,ne哪,ne讷,nei内,nei馁,nen嫩,nen恁,neng能,ng嗯,ni你,ni泥,ni拟,ni腻,ni逆,ni呢,ni溺,ni倪,ni尼,ni匿,ni妮,ni霓,ni铌,ni昵,ni坭,ni祢,ni猊,ni伲,ni怩,ni鲵,ni睨,ni旎,ni慝,"
		Pinyin_Dict = Pinyin_Dict & "nian年,nian念,nian捻,nian撵,nian拈,nian碾,nian蔫,nian廿,nian黏,nian辇,nian鲇,nian鲶,nian埝,niang娘,niang酿,niao鸟,niao尿,niao袅,niao茑,niao脲,niao嬲,nie捏,nie镍,nie聂,nie孽,nie涅,nie镊,nie啮,nie陧,nie蘖,nie嗫,nie臬,nie蹑,nie颞,nie乜,nin您,ning拧,ning凝,ning宁,ning柠,"
		Pinyin_Dict = Pinyin_Dict & "ning狞,ning泞,ning佞,ning甯,ning咛,ning聍,niu牛,niu扭,niu纽,niu钮,niu拗,niu妞,niu狃,niu忸,nong弄,nong浓,nong农,nong脓,nong哝,nong侬,nou耨,nu怒,nu努,nu奴,nu孥,nu胬,nu驽,nu弩,nuan暖,nue虐,nue疟,nuo挪,nuo诺,nuo懦,nuo糯,nuo娜,nuo喏,nuo傩,nuo锘,nuo搦,"
		Pinyin_Dict = Pinyin_Dict & "nv女,nv衄,nv钕,nv恧,nve虐,nve疟,o哦,o喔,o噢,ou偶,ou呕,ou欧,ou藕,ou鸥,ou区,ou沤,ou殴,ou怄,ou瓯,ou讴,ou耦,pa怕,pa爬,pa趴,pa啪,pa耙,pa扒,pa帕,pa琶,pa筢,pa杷,pa葩,pai派,pai排,pai拍,pai牌,pai迫,pai徘,pai湃,pai哌,"
		Pinyin_Dict = Pinyin_Dict & "pai俳,pai蒎,pan盘,pan盼,pan判,pan攀,pan畔,pan潘,pan叛,pan磐,pan番,pan胖,pan襻,pan蟠,pan袢,pan泮,pan拚,pan爿,pan蹒,pang旁,pang胖,pang耪,pang庞,pang乓,pang膀,pang磅,pang滂,pang彷,pang逄,pang螃,pang镑,pao跑,pao抛,pao炮,pao泡,pao刨,pao袍,pao咆,pao狍,pao匏,"
		Pinyin_Dict = Pinyin_Dict & "pao庖,pao疱,pao脬,pei陪,pei配,pei赔,pei呸,pei胚,pei佩,pei培,pei沛,pei裴,pei旆,pei锫,pei帔,pei醅,pei霈,pei辔,pen喷,pen盆,pen湓,peng碰,peng捧,peng棚,peng砰,peng蓬,peng朋,peng彭,peng鹏,peng烹,peng硼,peng膨,peng抨,peng澎,peng篷,peng怦,peng堋,peng蟛,peng嘭,pi批,"
		Pinyin_Dict = Pinyin_Dict & "pi皮,pi披,pi匹,pi劈,pi辟,pi坯,pi屁,pi脾,pi僻,pi疲,pi痞,pi霹,pi琵,pi毗,pi啤,pi譬,pi砒,pi否,pi貔,pi丕,pi圮,pi媲,pi癖,pi仳,pi擗,pi郫,pi甓,pi枇,pi睥,pi蜱,pi鼙,pi邳,pi陂,pi铍,pi庀,pi罴,pi埤,pi纰,pi陴,pi淠,"
		Pinyin_Dict = Pinyin_Dict & "pi噼,pi蚍,pi裨,pi疋,pi芘,pian片,pian篇,pian骗,pian偏,pian便,pian扁,pian翩,pian缏,pian犏,pian骈,pian胼,pian蹁,pian谝,piao忄,piao票,piao飘,piao漂,piao瓢,piao朴,piao螵,piao嫖,piao瞟,piao殍,piao缥,piao嘌,piao骠,piao剽,pie瞥,pie撇,pie氕,pie苤,pie丿,pin品,pin贫,pin聘,"
		Pinyin_Dict = Pinyin_Dict & "pin拼,pin频,pin嫔,pin榀,pin姘,pin牝,pin颦,ping平,ping凭,ping瓶,ping评,ping屏,ping乒,ping萍,ping苹,ping坪,ping冯,ping娉,ping鲆,ping枰,ping俜,po破,po坡,po颇,po婆,po泼,po迫,po泊,po魄,po朴,po繁,po粕,po笸,po皤,po钋,po陂,po鄱,po攴,po叵,po珀,"
		Pinyin_Dict = Pinyin_Dict & "po钷,pou剖,pou掊,pou裒,pu扑,pu铺,pu谱,pu脯,pu仆,pu蒲,pu葡,pu朴,pu菩,pu曝,pu莆,pu瀑,pu埔,pu圃,pu浦,pu堡,pu普,pu暴,pu镨,pu噗,pu匍,pu溥,pu濮,pu氆,pu蹼,pu璞,pu镤,qi起,qi其,qi七,qi气,qi期,qi齐,qi器,qi妻,qi骑,"
		Pinyin_Dict = Pinyin_Dict & "qi汽,qi棋,qi奇,qi欺,qi漆,qi启,qi戚,qi柒,qi岂,qi砌,qi弃,qi泣,qi祁,qi凄,qi企,qi乞,qi契,qi歧,qi祈,qi栖,qi畦,qi脐,qi崎,qi稽,qi迄,qi缉,qi沏,qi讫,qi旗,qi祺,qi颀,qi骐,qi屺,qi岐,qi蹊,qi蕲,qi桤,qi憩,qi芪,qi荠,"
		Pinyin_Dict = Pinyin_Dict & "qi萋,qi芑,qi汔,qi亟,qi鳍,qi俟,qi槭,qi嘁,qi蛴,qi綦,qi亓,qi欹,qi琪,qi麒,qi琦,qi蜞,qi圻,qi杞,qi葺,qi碛,qi淇,qi耆,qi绮,qi綮,qia恰,qia卡,qia掐,qia洽,qia髂,qia袷,qia葜,qian前,qian钱,qian千,qian牵,qian浅,qian签,qian欠,qian铅,qian嵌,"
		Pinyin_Dict = Pinyin_Dict & "qian钎,qian迁,qian钳,qian乾,qian谴,qian谦,qian潜,qian歉,qian纤,qian扦,qian遣,qian黔,qian堑,qian仟,qian岍,qian钤,qian褰,qian箝,qian掮,qian搴,qian倩,qian慊,qian悭,qian愆,qian虔,qian芡,qian荨,qian缱,qian佥,qian芊,qian阡,qian肷,qian茜,qian椠,qian犍,qian骞,qian羟,qian赶,qiang强,qiang枪,"
		Pinyin_Dict = Pinyin_Dict & "qiang墙,qiang抢,qiang腔,qiang呛,qiang羌,qiang蔷,qiang蜣,qiang跄,qiang戗,qiang襁,qiang戕,qiang炝,qiang镪,qiang锵,qiang羟,qiang樯,qiang嫱,qiao桥,qiao瞧,qiao敲,qiao巧,qiao翘,qiao锹,qiao壳,qiao鞘,qiao撬,qiao悄,qiao俏,qiao窍,qiao雀,qiao乔,qiao侨,qiao峭,qiao橇,qiao樵,qiao荞,qiao跷,qiao硗,qiao憔,qiao谯,"
		Pinyin_Dict = Pinyin_Dict & "qiao鞒,qiao愀,qiao缲,qiao诮,qiao劁,qiao峤,qiao搞,qiao铫,qie切,qie且,qie怯,qie窃,qie茄,qie郄,qie趄,qie惬,qie锲,qie妾,qie箧,qie慊,qie伽,qie挈,qin亲,qin琴,qin侵,qin勤,qin擒,qin寝,qin秦,qin芹,qin沁,qin禽,qin钦,qin吣,qin覃,qin矜,qin衾,qin芩,qin廑,qin嗪,"
		Pinyin_Dict = Pinyin_Dict & "qin螓,qin噙,qin揿,qin檎,qin锓,qing请,qing轻,qing清,qing青,qing情,qing晴,qing氢,qing倾,qing庆,qing擎,qing顷,qing亲,qing卿,qing氰,qing圊,qing謦,qing檠,qing箐,qing苘,qing蜻,qing黥,qing罄,qing鲭,qing磬,qing綮,qiong穷,qiong琼,qiong跫,qiong穹,qiong邛,qiong蛩,qiong茕,qiong銎,qiong筇,qiu求,"
		Pinyin_Dict = Pinyin_Dict & "qiu球,qiu秋,qiu丘,qiu泅,qiu仇,qiu邱,qiu囚,qiu酋,qiu龟,qiu楸,qiu蚯,qiu裘,qiu糗,qiu蝤,qiu巯,qiu逑,qiu俅,qiu虬,qiu赇,qiu鳅,qiu犰,qiu湫,qiu遒,qu去,qu取,qu区,qu娶,qu渠,qu曲,qu趋,qu趣,qu屈,qu驱,qu蛆,qu躯,qu龋,qu戌,qu蠼,qu蘧,qu祛,"
		Pinyin_Dict = Pinyin_Dict & "qu蕖,qu磲,qu劬,qu诎,qu鸲,qu阒,qu麴,qu癯,qu衢,qu黢,qu璩,qu氍,qu觑,qu蛐,qu朐,qu瞿,qu岖,qu苣,quan全,quan权,quan劝,quan圈,quan拳,quan犬,quan泉,quan券,quan颧,quan痊,quan醛,quan铨,quan筌,quan绻,quan诠,quan辁,quan畎,quan鬈,quan悛,quan蜷,quan荃,quan犭,"
		Pinyin_Dict = Pinyin_Dict & "que却,que缺,que确,que雀,que瘸,que鹊,que炔,que榷,que阙,que阕,que悫,qun群,qun裙,qun麇,qun逡,ran染,ran燃,ran然,ran冉,ran髯,ran苒,ran蚺,rang让,rang嚷,rang瓤,rang攘,rang壤,rang穰,rang禳,rao饶,rao绕,rao扰,rao荛,rao桡,rao娆,re热,re惹,re喏,ren人,ren任,"
		Pinyin_Dict = Pinyin_Dict & "ren忍,ren认,ren刃,ren仁,ren韧,ren妊,ren纫,ren壬,ren饪,ren轫,ren仞,ren荏,ren葚,ren衽,ren稔,ren亻,reng仍,reng扔,ri日,rong容,rong绒,rong融,rong溶,rong熔,rong荣,rong戎,rong蓉,rong冗,rong茸,rong榕,rong狨,rong嵘,rong肜,rong蝾,rou肉,rou揉,rou柔,rou糅,rou蹂,rou鞣,"
		Pinyin_Dict = Pinyin_Dict & "ru如,ru入,ru汝,ru儒,ru茹,ru乳,ru褥,ru辱,ru蠕,ru孺,ru蓐,ru襦,ru铷,ru嚅,ru缛,ru濡,ru薷,ru颥,ru溽,ru洳,ruan软,ruan阮,ruan朊,rui瑞,rui蕊,rui锐,rui睿,rui芮,rui蚋,rui枘,rui蕤,run润,run闰,ruo若,ruo弱,ruo箬,ruo偌,sa撒,sa洒,sa萨,"
		Pinyin_Dict = Pinyin_Dict & "sa仨,sa卅,sa飒,sa脎,sai塞,sai腮,sai鳃,sai赛,sai噻,san三,san散,san伞,san叁,san馓,san糁,san毵,sang桑,sang丧,sang嗓,sang颡,sang磉,sang搡,sao扫,sao嫂,sao搔,sao骚,sao埽,sao鳋,sao臊,sao缫,sao瘙,se色,se涩,se瑟,se塞,se啬,se铯,se穑,sen森,seng僧,"
		Pinyin_Dict = Pinyin_Dict & "sha杀,sha沙,sha啥,sha纱,sha傻,sha砂,sha刹,sha莎,sha厦,sha煞,sha杉,sha唼,sha鲨,sha霎,sha铩,sha痧,sha裟,sha歃,shai晒,shai筛,shai色,shan山,shan闪,shan衫,shan善,shan扇,shan杉,shan删,shan煽,shan单,shan珊,shan掺,shan赡,shan栅,shan苫,shan膳,shan陕,shan汕,shan擅,shan缮,"
		Pinyin_Dict = Pinyin_Dict & "shan嬗,shan蟮,shan芟,shan禅,shan跚,shan鄯,shan潸,shan鳝,shan姗,shan剡,shan骟,shan疝,shan膻,shan讪,shan钐,shan舢,shan埏,shan彡,shan髟,shang上,shang伤,shang尚,shang商,shang赏,shang晌,shang墒,shang裳,shang熵,shang觞,shang绱,shang殇,shang垧,shao少,shao烧,shao捎,shao哨,shao勺,shao梢,shao稍,shao邵,"
		Pinyin_Dict = Pinyin_Dict & "shao韶,shao绍,shao芍,shao鞘,shao苕,shao劭,shao潲,shao艄,shao蛸,shao筲,she社,she射,she蛇,she设,she舌,she摄,she舍,she折,she涉,she赊,she赦,she慑,she奢,she歙,she厍,she畲,she猞,she麝,she滠,she佘,she揲,shei谁,shen身,shen伸,shen深,shen婶,shen神,shen甚,shen渗,shen肾,"
		Pinyin_Dict = Pinyin_Dict & "shen审,shen申,shen沈,shen绅,shen呻,shen参,shen砷,shen什,shen娠,shen慎,shen葚,shen莘,shen诜,shen谂,shen矧,shen椹,shen渖,shen蜃,shen哂,shen胂,sheng声,sheng省,sheng剩,sheng生,sheng升,sheng绳,sheng胜,sheng盛,sheng圣,sheng甥,sheng牲,sheng乘,sheng晟,sheng渑,sheng眚,sheng笙,sheng嵊,shi是,shi使,shi十,"
		Pinyin_Dict = Pinyin_Dict & "shi时,shi事,shi室,shi市,shi石,shi师,shi试,shi史,shi式,shi识,shi虱,shi矢,shi拾,shi屎,shi驶,shi始,shi似,shi示,shi士,shi世,shi柿,shi匙,shi拭,shi誓,shi逝,shi势,shi什,shi殖,shi峙,shi嗜,shi噬,shi失,shi适,shi仕,shi侍,shi释,shi饰,shi氏,shi狮,shi食,"
		Pinyin_Dict = Pinyin_Dict & "shi恃,shi蚀,shi视,shi实,shi施,shi湿,shi诗,shi尸,shi豕,shi莳,shi埘,shi铈,shi舐,shi鲥,shi鲺,shi贳,shi轼,shi蓍,shi筮,shi炻,shi谥,shi弑,shi酾,shi螫,shi礻,shi铊,shi饣,shou手,shou受,shou收,shou首,shou守,shou瘦,shou授,shou兽,shou售,shou寿,shou艏,shou狩,shou绶,"
		Pinyin_Dict = Pinyin_Dict & "shou扌,shu书,shu树,shu数,shu熟,shu输,shu梳,shu叔,shu属,shu束,shu术,shu述,shu蜀,shu黍,shu鼠,shu淑,shu赎,shu孰,shu蔬,shu疏,shu戍,shu竖,shu墅,shu庶,shu薯,shu漱,shu恕,shu枢,shu暑,shu殊,shu抒,shu曙,shu署,shu舒,shu姝,shu摅,shu秫,shu纾,shu沭,shu毹,"
		Pinyin_Dict = Pinyin_Dict & "shu腧,shu塾,shu菽,shu殳,shu澍,shu倏,shu疋,shu镯,shua刷,shua耍,shua唰,shuai摔,shuai甩,shuai率,shuai帅,shuai衰,shuai蟀,shuan栓,shuan拴,shuan闩,shuan涮,shuang双,shuang霜,shuang爽,shuang泷,shuang孀,shui水,shui睡,shui税,shui说,shui氵,shun顺,shun吮,shun瞬,shun舜,shuo说,shuo数,shuo硕,shuo烁,shuo朔,"
		Pinyin_Dict = Pinyin_Dict & "shuo搠,shuo妁,shuo槊,shuo蒴,shuo铄,si四,si死,si丝,si撕,si似,si私,si嘶,si思,si寺,si司,si斯,si伺,si肆,si饲,si嗣,si巳,si耜,si驷,si兕,si蛳,si厮,si汜,si锶,si泗,si笥,si咝,si鸶,si姒,si厶,si缌,si祀,si澌,si俟,si徙,song送,"
		Pinyin_Dict = Pinyin_Dict & "song松,song耸,song宋,song颂,song诵,song怂,song讼,song竦,song菘,song淞,song悚,song嵩,song凇,song崧,song忪,sou艘,sou搜,sou擞,sou嗽,sou嗾,sou嗖,sou飕,sou叟,sou薮,sou锼,sou馊,sou瞍,sou溲,sou螋,su素,su速,su诉,su塑,su宿,su俗,su苏,su肃,su粟,su酥,su缩,"
		Pinyin_Dict = Pinyin_Dict & "su溯,su僳,su愫,su簌,su觫,su稣,su夙,su嗉,su谡,su蔌,su涑,suan酸,suan算,suan蒜,suan狻,sui岁,sui随,sui碎,sui虽,sui穗,sui遂,sui尿,sui隋,sui髓,sui绥,sui隧,sui祟,sui眭,sui谇,sui濉,sui邃,sui燧,sui荽,sui睢,sun孙,sun损,sun笋,sun榫,sun荪,sun飧,"
		Pinyin_Dict = Pinyin_Dict & "sun狲,sun隼,suo所,suo缩,suo锁,suo琐,suo索,suo梭,suo蓑,suo莎,suo唆,suo挲,suo睃,suo嗍,suo唢,suo桫,suo嗦,suo娑,suo羧,ta他,ta她,ta它,ta踏,ta塔,ta塌,ta拓,ta獭,ta挞,ta蹋,ta溻,ta趿,ta鳎,ta沓,ta榻,ta漯,ta遢,ta铊,ta闼,tai太,tai抬,"
		Pinyin_Dict = Pinyin_Dict & "tai台,tai态,tai胎,tai苔,tai泰,tai酞,tai汰,tai炱,tai肽,tai跆,tai鲐,tai钛,tai薹,tai邰,tai骀,tan谈,tan叹,tan探,tan滩,tan弹,tan碳,tan摊,tan潭,tan贪,tan坛,tan痰,tan毯,tan坦,tan炭,tan瘫,tan谭,tan坍,tan檀,tan袒,tan钽,tan郯,tan镡,tan锬,tan覃,tan澹,"
		Pinyin_Dict = Pinyin_Dict & "tan昙,tan忐,tan赕,tang躺,tang趟,tang堂,tang糖,tang汤,tang塘,tang烫,tang倘,tang淌,tang唐,tang搪,tang棠,tang膛,tang螳,tang樘,tang羰,tang醣,tang瑭,tang镗,tang傥,tang饧,tang溏,tang耥,tang帑,tang铴,tang螗,tang铛,tao套,tao掏,tao逃,tao桃,tao讨,tao淘,tao涛,tao滔,tao陶,tao绦,"
		Pinyin_Dict = Pinyin_Dict & "tao萄,tao鼗,tao洮,tao焘,tao啕,tao饕,tao韬,tao叨,te特,te铽,te忑,te忒,teng疼,teng腾,teng藤,teng誊,teng滕,ti提,ti替,ti体,ti题,ti踢,ti蹄,ti剃,ti剔,ti梯,ti锑,ti啼,ti涕,ti嚏,ti惕,ti屉,ti醍,ti鹈,ti绨,ti缇,ti倜,ti裼,ti逖,ti荑,"
		Pinyin_Dict = Pinyin_Dict & "ti悌,tian天,tian田,tian添,tian填,tian甜,tian舔,tian恬,tian腆,tian掭,tian钿,tian阗,tian忝,tian殄,tian畋,tian锘,tiao条,tiao跳,tiao挑,tiao调,tiao迢,tiao眺,tiao龆,tiao笤,tiao祧,tiao蜩,tiao髫,tiao佻,tiao窕,tiao鲦,tiao苕,tiao粜,tiao铫,tie铁,tie贴,tie帖,tie萜,tie餮,tie锇,ting听,"
		Pinyin_Dict = Pinyin_Dict & "ting停,ting挺,ting厅,ting亭,ting艇,ting庭,ting廷,ting烃,ting汀,ting莛,ting铤,ting葶,ting婷,ting蜓,ting梃,ting霆,tong同,tong通,tong痛,tong铜,tong桶,tong筒,tong捅,tong统,tong童,tong彤,tong桐,tong瞳,tong酮,tong潼,tong茼,tong仝,tong砼,tong峒,tong恸,tong佟,tong嗵,tong垌,tong僮,tou头,"
		Pinyin_Dict = Pinyin_Dict & "tou偷,tou透,tou投,tou钭,tou骰,tou亠,tu土,tu图,tu兔,tu涂,tu吐,tu秃,tu突,tu徒,tu凸,tu途,tu屠,tu酴,tu荼,tu钍,tu菟,tu堍,tuan团,tuan湍,tuan疃,tuan抟,tuan彖,tui腿,tui推,tui退,tui褪,tui颓,tui蜕,tui煺,tun吞,tun屯,tun褪,tun臀,tun囤,tun氽,"
		Pinyin_Dict = Pinyin_Dict & "tun饨,tun豚,tun暾,tuo拖,tuo脱,tuo托,tuo妥,tuo驮,tuo拓,tuo驼,tuo椭,tuo唾,tuo鸵,tuo陀,tuo橐,tuo柝,tuo跎,tuo乇,tuo坨,tuo佗,tuo庹,tuo酡,tuo柁,tuo鼍,tuo沱,tuo箨,tuo砣,tuo说,tuo铊,wa挖,wa瓦,wa蛙,wa哇,wa娃,wa洼,wa袜,wa佤,wa娲,wa腽,wai外,"
		Pinyin_Dict = Pinyin_Dict & "wai歪,wai崴,wan完,wan万,wan晚,wan碗,wan玩,wan弯,wan挽,wan湾,wan丸,wan腕,wan宛,wan婉,wan烷,wan顽,wan豌,wan惋,wan皖,wan蔓,wan莞,wan脘,wan蜿,wan绾,wan芄,wan琬,wan纨,wan剜,wan畹,wan菀,wang望,wang忘,wang王,wang往,wang网,wang亡,wang枉,wang旺,wang汪,wang妄,"
		Pinyin_Dict = Pinyin_Dict & "wang辋,wang魍,wang惘,wang罔,wang尢,wei为,wei位,wei未,wei围,wei喂,wei胃,wei微,wei味,wei尾,wei伪,wei威,wei伟,wei卫,wei危,wei违,wei委,wei魏,wei唯,wei维,wei畏,wei惟,wei韦,wei巍,wei蔚,wei谓,wei尉,wei潍,wei纬,wei慰,wei桅,wei萎,wei苇,wei渭,wei葳,wei帏,"
		Pinyin_Dict = Pinyin_Dict & "wei艉,wei鲔,wei娓,wei逶,wei闱,wei隈,wei沩,wei玮,wei涠,wei帷,wei崴,wei隗,wei诿,wei洧,wei偎,wei猥,wei猬,wei嵬,wei軎,wei韪,wei炜,wei煨,wei圩,wei薇,wei痿,wei囗,wen问,wen文,wen闻,wen稳,wen温,wen吻,wen蚊,wen纹,wen瘟,wen紊,wen汶,wen阌,wen刎,wen雯,"
		Pinyin_Dict = Pinyin_Dict & "wen璺,weng翁,weng嗡,weng瓮,weng蕹,weng蓊,wo我,wo握,wo窝,wo卧,wo挝,wo沃,wo蜗,wo涡,wo斡,wo倭,wo幄,wo龌,wo肟,wo莴,wo喔,wo渥,wo硪,wu无,wu五,wu屋,wu物,wu舞,wu雾,wu误,wu捂,wu污,wu悟,wu勿,wu钨,wu武,wu戊,wu务,wu呜,wu伍,"
		Pinyin_Dict = Pinyin_Dict & "wu吴,wu午,wu吾,wu侮,wu乌,wu毋,wu恶,wu诬,wu芜,wu巫,wu晤,wu梧,wu坞,wu妩,wu蜈,wu牾,wu寤,wu兀,wu怃,wu阢,wu邬,wu唔,wu忤,wu骛,wu於,wu鋈,wu仵,wu杌,wu鹜,wu婺,wu迕,wu痦,wu芴,wu焐,wu庑,wu鹉,wu鼯,wu浯,wu圬,xi西,"
		Pinyin_Dict = Pinyin_Dict & "xi洗,xi细,xi吸,xi戏,xi系,xi喜,xi席,xi稀,xi溪,xi熄,xi锡,xi膝,xi息,xi袭,xi惜,xi习,xi嘻,xi夕,xi悉,xi矽,xi熙,xi希,xi檄,xi牺,xi晰,xi昔,xi媳,xi硒,xi铣,xi烯,xi析,xi隙,xi汐,xi犀,xi蜥,xi奚,xi浠,xi葸,xi饩,xi屣,"
		Pinyin_Dict = Pinyin_Dict & "xi玺,xi嬉,xi禊,xi兮,xi翕,xi穸,xi禧,xi僖,xi淅,xi蓰,xi舾,xi蹊,xi醯,xi郗,xi欷,xi皙,xi蟋,xi羲,xi茜,xi徙,xi隰,xi唏,xi曦,xi螅,xi歙,xi樨,xi阋,xi粞,xi熹,xi觋,xi菥,xi鼷,xi裼,xi舄,xia下,xia吓,xia夏,xia峡,xia虾,xia瞎,"
		Pinyin_Dict = Pinyin_Dict & "xia霞,xia狭,xia匣,xia侠,xia辖,xia厦,xia暇,xia狎,xia柙,xia呷,xia黠,xia硖,xia罅,xia遐,xia瑕,xian先,xian线,xian县,xian现,xian显,xian掀,xian闲,xian献,xian嫌,xian陷,xian险,xian鲜,xian弦,xian衔,xian馅,xian限,xian咸,xian锨,xian仙,xian腺,xian贤,xian纤,xian宪,xian舷,xian涎,"
		Pinyin_Dict = Pinyin_Dict & "xian羡,xian铣,xian苋,xian藓,xian岘,xian痫,xian莶,xian籼,xian娴,xian蚬,xian猃,xian祆,xian冼,xian燹,xian跣,xian跹,xian酰,xian暹,xian氙,xian鹇,xian筅,xian霰,xian洗,xiang想,xiang向,xiang象,xiang项,xiang响,xiang香,xiang乡,xiang相,xiang像,xiang箱,xiang巷,xiang享,xiang镶,xiang厢,xiang降,xiang翔,xiang祥,"
		Pinyin_Dict = Pinyin_Dict & "xiang橡,xiang详,xiang湘,xiang襄,xiang飨,xiang鲞,xiang骧,xiang蟓,xiang庠,xiang芗,xiang饷,xiang缃,xiang葙,xiao小,xiao笑,xiao消,xiao削,xiao销,xiao萧,xiao效,xiao宵,xiao晓,xiao肖,xiao孝,xiao硝,xiao淆,xiao啸,xiao霄,xiao哮,xiao嚣,xiao校,xiao魈,xiao蛸,xiao骁,xiao枵,xiao哓,xiao筱,xiao潇,xiao逍,xiao枭,"
		Pinyin_Dict = Pinyin_Dict & "xiao绡,xiao箫,xie写,xie些,xie鞋,xie歇,xie斜,xie血,xie谢,xie卸,xie挟,xie屑,xie蟹,xie泻,xie懈,xie泄,xie楔,xie邪,xie协,xie械,xie谐,xie蝎,xie携,xie胁,xie解,xie叶,xie绁,xie颉,xie缬,xie獬,xie榭,xie廨,xie撷,xie偕,xie瀣,xie渫,xie亵,xie榍,xie邂,xie薤,"
		Pinyin_Dict = Pinyin_Dict & "xie躞,xie燮,xie勰,xie骱,xie鲑,xin新,xin心,xin欣,xin信,xin芯,xin薪,xin锌,xin辛,xin衅,xin忻,xin歆,xin囟,xin莘,xin镡,xin馨,xin鑫,xin昕,xin忄,xing性,xing行,xing型,xing形,xing星,xing醒,xing姓,xing腥,xing刑,xing杏,xing兴,xing幸,xing邢,xing猩,xing惺,xing省,xing硎,"
		Pinyin_Dict = Pinyin_Dict & "xing悻,xing荥,xing陉,xing擤,xing荇,xing研,xing饧,xiong胸,xiong雄,xiong凶,xiong兄,xiong熊,xiong汹,xiong匈,xiong芎,xiu修,xiu锈,xiu绣,xiu休,xiu羞,xiu宿,xiu嗅,xiu袖,xiu秀,xiu朽,xiu臭,xiu溴,xiu貅,xiu馐,xiu髹,xiu鸺,xiu咻,xiu庥,xiu岫,xu许,xu须,xu需,xu虚,xu嘘,xu蓄,"
		Pinyin_Dict = Pinyin_Dict & "xu续,xu序,xu叙,xu畜,xu絮,xu婿,xu戌,xu徐,xu旭,xu绪,xu吁,xu酗,xu恤,xu墟,xu糈,xu勖,xu栩,xu浒,xu蓿,xu顼,xu圩,xu洫,xu胥,xu醑,xu诩,xu溆,xu煦,xu盱,xuan选,xuan悬,xuan旋,xuan玄,xuan宣,xuan喧,xuan轩,xuan绚,xuan眩,xuan癣,xuan券,xuan暄,"
		Pinyin_Dict = Pinyin_Dict & "xuan楦,xuan儇,xuan渲,xuan漩,xuan泫,xuan铉,xuan璇,xuan煊,xuan碹,xuan镟,xuan炫,xuan揎,xuan萱,xuan谖,xue学,xue雪,xue血,xue靴,xue穴,xue削,xue薛,xue踅,xue噱,xue鳕,xue泶,xue谑,xun寻,xun讯,xun熏,xun训,xun循,xun殉,xun旬,xun巡,xun迅,xun驯,xun汛,xun逊,xun勋,xun询,"
		Pinyin_Dict = Pinyin_Dict & "xun浚,xun巽,xun鲟,xun浔,xun埙,xun恂,xun獯,xun醺,xun洵,xun郇,xun峋,xun蕈,xun薰,xun荀,xun窨,xun曛,xun徇,xun荨,ya呀,ya压,ya牙,ya押,ya芽,ya鸭,ya轧,ya崖,ya哑,ya亚,ya涯,ya丫,ya雅,ya衙,ya鸦,ya讶,ya蚜,ya垭,ya疋,ya砑,ya琊,ya桠,"
		Pinyin_Dict = Pinyin_Dict & "ya睚,ya娅,ya痖,ya岈,ya氩,ya伢,ya迓,ya揠,yan眼,yan烟,yan沿,yan盐,yan言,yan演,yan严,yan咽,yan淹,yan炎,yan掩,yan厌,yan宴,yan岩,yan研,yan延,yan堰,yan验,yan艳,yan殷,yan阉,yan砚,yan雁,yan唁,yan彦,yan焰,yan蜒,yan衍,yan谚,yan燕,yan颜,yan阎,"
		Pinyin_Dict = Pinyin_Dict & "yan铅,yan焉,yan奄,yan芫,yan厣,yan阏,yan菸,yan魇,yan琰,yan滟,yan焱,yan赝,yan筵,yan腌,yan兖,yan剡,yan餍,yan恹,yan罨,yan檐,yan湮,yan偃,yan谳,yan胭,yan晏,yan闫,yan俨,yan郾,yan酽,yan鄢,yan妍,yan鼹,yan崦,yan阽,yan嫣,yan涎,yan讠,yang样,yang养,yang羊,"
		Pinyin_Dict = Pinyin_Dict & "yang洋,yang仰,yang扬,yang秧,yang氧,yang痒,yang杨,yang漾,yang阳,yang殃,yang央,yang鸯,yang佯,yang疡,yang炀,yang恙,yang徉,yang鞅,yang泱,yang蛘,yang烊,yang怏,yao要,yao摇,yao药,yao咬,yao腰,yao窑,yao舀,yao邀,yao妖,yao谣,yao遥,yao姚,yao瑶,yao耀,yao尧,yao钥,yao侥,yao疟,"
		Pinyin_Dict = Pinyin_Dict & "yao珧,yao夭,yao鳐,yao鹞,yao轺,yao爻,yao吆,yao铫,yao幺,yao崾,yao肴,yao曜,yao徭,yao杳,yao窈,yao啮,yao繇,ye也,ye夜,ye业,ye野,ye叶,ye爷,ye页,ye液,ye掖,ye腋,ye冶,ye噎,ye耶,ye咽,ye曳,ye椰,ye邪,ye谒,ye邺,ye晔,ye烨,ye揶,ye铘,"
		Pinyin_Dict = Pinyin_Dict & "ye靥,yi一,yi以,yi已,yi亿,yi衣,yi移,yi依,yi易,yi医,yi乙,yi仪,yi亦,yi椅,yi益,yi倚,yi姨,yi翼,yi译,yi伊,yi遗,yi艾,yi胰,yi疑,yi沂,yi宜,yi异,yi彝,yi壹,yi蚁,yi谊,yi揖,yi铱,yi矣,yi翌,yi艺,yi抑,yi绎,yi邑,yi屹,"
		Pinyin_Dict = Pinyin_Dict & "yi尾,yi役,yi臆,yi逸,yi肄,yi疫,yi颐,yi裔,yi意,yi毅,yi忆,yi义,yi夷,yi溢,yi诣,yi议,yi怿,yi痍,yi镒,yi癔,yi怡,yi驿,yi旖,yi熠,yi酏,yi翊,yi欹,yi峄,yi圯,yi殪,yi咦,yi懿,yi噫,yi劓,yi诒,yi饴,yi漪,yi佚,yi咿,yi瘗,"
		Pinyin_Dict = Pinyin_Dict & "yi猗,yi眙,yi羿,yi弈,yi苡,yi荑,yi佾,yi贻,yi钇,yi缢,yi迤,yi刈,yi悒,yi黟,yi翳,yi弋,yi奕,yi蜴,yi埸,yi挹,yi嶷,yi薏,yi呓,yi轶,yi镱,yi舣,yi奇,yi硪,yi衤,yi铊,yin因,yin引,yin印,yin银,yin音,yin饮,yin阴,yin隐,yin荫,yin吟,"
		Pinyin_Dict = Pinyin_Dict & "yin尹,yin寅,yin茵,yin淫,yin殷,yin姻,yin堙,yin鄞,yin喑,yin夤,yin胤,yin龈,yin吲,yin狺,yin垠,yin霪,yin蚓,yin氤,yin铟,yin窨,yin瘾,yin洇,yin茚,yin廴,ying应,ying硬,ying影,ying营,ying迎,ying映,ying蝇,ying赢,ying鹰,ying英,ying颖,ying莹,ying盈,ying婴,ying樱,ying缨,"
		Pinyin_Dict = Pinyin_Dict & "ying荧,ying萤,ying萦,ying楹,ying蓥,ying瘿,ying茔,ying鹦,ying媵,ying莺,ying璎,ying郢,ying嘤,ying撄,ying瑛,ying滢,ying潆,ying嬴,ying罂,ying瀛,ying膺,ying荥,ying颍,yo哟,yo育,yo唷,yong用,yong涌,yong永,yong拥,yong蛹,yong勇,yong雍,yong咏,yong泳,yong佣,yong踊,yong痈,yong庸,yong臃,"
		Pinyin_Dict = Pinyin_Dict & "yong恿,yong壅,yong慵,yong俑,yong墉,yong鳙,yong邕,yong喁,yong甬,yong饔,yong镛,you有,you又,you由,you右,you油,you游,you幼,you优,you友,you铀,you忧,you尤,you犹,you诱,you悠,you邮,you酉,you佑,you釉,you幽,you疣,you攸,you蚰,you莠,you鱿,you卣,you黝,you莸,you猷,"
		Pinyin_Dict = Pinyin_Dict & "you蚴,you宥,you牖,you囿,you柚,you蝣,you莜,you鼬,you铕,you蝤,you繇,you呦,you侑,you尢,yu与,yu于,yu欲,yu鱼,yu雨,yu余,yu遇,yu语,yu愈,yu狱,yu玉,yu渔,yu予,yu誉,yu育,yu愚,yu羽,yu虞,yu娱,yu淤,yu舆,yu屿,yu禹,yu宇,yu迂,yu俞,"
		Pinyin_Dict = Pinyin_Dict & "yu逾,yu域,yu芋,yu郁,yu吁,yu盂,yu喻,yu峪,yu御,yu愉,yu渝,yu尉,yu榆,yu隅,yu浴,yu寓,yu裕,yu预,yu豫,yu驭,yu蔚,yu妪,yu嵛,yu雩,yu馀,yu阈,yu窬,yu鹆,yu妤,yu揄,yu窳,yu觎,yu臾,yu舁,yu龉,yu蓣,yu煜,yu钰,yu谀,yu纡,"
		Pinyin_Dict = Pinyin_Dict & "yu於,yu竽,yu瑜,yu禺,yu聿,yu欤,yu俣,yu伛,yu圄,yu鹬,yu庾,yu昱,yu萸,yu瘐,yu谕,yu鬻,yu圉,yu瘀,yu熨,yu饫,yu毓,yu燠,yu腴,yu狳,yu菀,yu蜮,yu蝓,yu吾,yuan远,yuan员,yuan元,yuan院,yuan圆,yuan原,yuan愿,yuan园,yuan援,yuan猿,yuan怨,yuan冤,"
		Pinyin_Dict = Pinyin_Dict & "yuan源,yuan缘,yuan袁,yuan渊,yuan苑,yuan垣,yuan鸳,yuan辕,yuan圜,yuan鼋,yuan橼,yuan媛,yuan爰,yuan眢,yuan鸢,yuan掾,yuan芫,yuan沅,yuan瑗,yuan螈,yuan箢,yuan塬,yue月,yue越,yue约,yue跃,yue阅,yue乐,yue岳,yue悦,yue曰,yue说,yue粤,yue钥,yue瀹,yue钺,yue刖,yue龠,yue栎,yue樾,"
		Pinyin_Dict = Pinyin_Dict & "yue哕,yun云,yun运,yun晕,yun允,yun匀,yun韵,yun陨,yun孕,yun耘,yun蕴,yun酝,yun郧,yun员,yun氲,yun恽,yun愠,yun郓,yun芸,yun筠,yun韫,yun昀,yun狁,yun殒,yun纭,yun熨,za杂,za砸,za咋,za匝,za扎,za咂,za拶,zai在,zai再,zai灾,zai载,zai栽,zai宰,zai哉,"
		Pinyin_Dict = Pinyin_Dict & "zai甾,zai崽,zan咱,zan暂,zan攒,zan赞,zan簪,zan趱,zan糌,zan瓒,zan拶,zan昝,zan錾,zang脏,zang葬,zang赃,zang藏,zang臧,zang驵,zao早,zao造,zao遭,zao糟,zao灶,zao燥,zao枣,zao凿,zao躁,zao藻,zao皂,zao噪,zao澡,zao蚤,zao唣,ze则,ze责,ze择,ze泽,ze咋,ze箦,"
		Pinyin_Dict = Pinyin_Dict & "ze舴,ze帻,ze迮,ze啧,ze仄,ze昃,ze笮,ze赜,zei贼,zen怎,zen谮,zeng增,zeng赠,zeng憎,zeng曾,zeng缯,zeng罾,zeng甑,zeng锃,zha扎,zha炸,zha渣,zha闸,zha眨,zha榨,zha乍,zha轧,zha诈,zha铡,zha札,zha查,zha栅,zha咋,zha喳,zha砟,zha痄,zha吒,zha哳,zha楂,zha蚱,"
		Pinyin_Dict = Pinyin_Dict & "zha揸,zha喋,zha柞,zha咤,zha齄,zha龃,zhai摘,zhai窄,zhai债,zhai斋,zhai寨,zhai择,zhai翟,zhai宅,zhai砦,zhai瘵,zhan站,zhan占,zhan战,zhan盏,zhan沾,zhan粘,zhan毡,zhan展,zhan栈,zhan詹,zhan颤,zhan蘸,zhan湛,zhan绽,zhan斩,zhan辗,zhan崭,zhan瞻,zhan谵,zhan搌,zhan旃,zhan骣,zhang张,zhang章,"
		Pinyin_Dict = Pinyin_Dict & "zhang长,zhang帐,zhang仗,zhang丈,zhang掌,zhang涨,zhang账,zhang樟,zhang杖,zhang彰,zhang漳,zhang胀,zhang瘴,zhang障,zhang仉,zhang嫜,zhang幛,zhang鄣,zhang璋,zhang嶂,zhang獐,zhang蟑,zhao找,zhao着,zhao照,zhao招,zhao罩,zhao爪,zhao兆,zhao朝,zhao昭,zhao沼,zhao肇,zhao召,zhao赵,zhao棹,zhao啁,zhao钊,zhao笊,zhao诏,"
		Pinyin_Dict = Pinyin_Dict & "zhe着,zhe这,zhe者,zhe折,zhe遮,zhe蛰,zhe哲,zhe蔗,zhe锗,zhe辙,zhe浙,zhe柘,zhe辄,zhe赭,zhe摺,zhe鹧,zhe磔,zhe褶,zhe蜇,zhe谪,zhen真,zhen阵,zhen镇,zhen针,zhen震,zhen枕,zhen振,zhen斟,zhen珍,zhen疹,zhen诊,zhen甄,zhen砧,zhen臻,zhen贞,zhen侦,zhen缜,zhen蓁,zhen祯,zhen箴,"
		Pinyin_Dict = Pinyin_Dict & "zhen轸,zhen榛,zhen稹,zhen赈,zhen朕,zhen鸩,zhen胗,zhen浈,zhen桢,zhen畛,zhen圳,zhen椹,zhen溱,zheng正,zheng整,zheng睁,zheng争,zheng挣,zheng征,zheng怔,zheng证,zheng症,zheng郑,zheng拯,zheng蒸,zheng狰,zheng政,zheng峥,zheng钲,zheng铮,zheng筝,zheng诤,zheng徵,zheng鲭,zhi只,zhi之,zhi直,zhi知,zhi制,zhi指,"
		Pinyin_Dict = Pinyin_Dict & "zhi纸,zhi支,zhi芝,zhi枝,zhi稚,zhi吱,zhi蜘,zhi质,zhi肢,zhi脂,zhi汁,zhi炙,zhi织,zhi职,zhi痔,zhi植,zhi抵,zhi殖,zhi执,zhi值,zhi侄,zhi址,zhi滞,zhi止,zhi趾,zhi治,zhi旨,zhi窒,zhi志,zhi挚,zhi掷,zhi至,zhi致,zhi置,zhi帜,zhi识,zhi峙,zhi智,zhi秩,zhi帙,"
		Pinyin_Dict = Pinyin_Dict & "zhi摭,zhi黹,zhi桎,zhi枳,zhi轵,zhi忮,zhi祉,zhi蛭,zhi膣,zhi觯,zhi郅,zhi栀,zhi彘,zhi芷,zhi祗,zhi咫,zhi鸷,zhi絷,zhi踬,zhi胝,zhi骘,zhi轾,zhi痣,zhi陟,zhi踯,zhi雉,zhi埴,zhi贽,zhi卮,zhi酯,zhi豸,zhi跖,zhi栉,zhi夂,zhi徵,zhong中,zhong重,zhong种,zhong钟,zhong肿,"
		Pinyin_Dict = Pinyin_Dict & "zhong众,zhong终,zhong盅,zhong忠,zhong仲,zhong衷,zhong踵,zhong舯,zhong螽,zhong锺,zhong冢,zhong忪,zhou周,zhou洲,zhou皱,zhou粥,zhou州,zhou轴,zhou舟,zhou昼,zhou骤,zhou宙,zhou诌,zhou肘,zhou帚,zhou咒,zhou繇,zhou胄,zhou纣,zhou荮,zhou啁,zhou碡,zhou绉,zhou籀,zhou妯,zhou酎,zhu住,zhu主,zhu猪,zhu竹,"
		Pinyin_Dict = Pinyin_Dict & "zhu株,zhu煮,zhu筑,zhu著,zhu贮,zhu铸,zhu嘱,zhu拄,zhu注,zhu祝,zhu驻,zhu属,zhu术,zhu珠,zhu瞩,zhu蛛,zhu朱,zhu柱,zhu诸,zhu诛,zhu逐,zhu助,zhu烛,zhu蛀,zhu潴,zhu洙,zhu伫,zhu瘃,zhu翥,zhu茱,zhu苎,zhu橥,zhu舳,zhu杼,zhu箸,zhu炷,zhu侏,zhu铢,zhu疰,zhu渚,"
		Pinyin_Dict = Pinyin_Dict & "zhu褚,zhu躅,zhu麈,zhu邾,zhu槠,zhu竺,zhu丶,zhua抓,zhua爪,zhua挝,zhuai拽,zhuai转,zhuan转,zhuan专,zhuan砖,zhuan赚,zhuan传,zhuan撰,zhuan篆,zhuan颛,zhuan馔,zhuan啭,zhuan沌,zhuang装,zhuang撞,zhuang庄,zhuang壮,zhuang桩,zhuang状,zhuang幢,zhuang妆,zhuang奘,zhuang戆,zhui追,zhui坠,zhui缀,zhui锥,zhui赘,zhui椎,zhui骓,"
		Pinyin_Dict = Pinyin_Dict & "zhui惴,zhui缒,zhui隹,zhun准,zhun谆,zhun肫,zhun窀,zhun饨,zhuo捉,zhuo桌,zhuo着,zhuo啄,zhuo拙,zhuo灼,zhuo浊,zhuo卓,zhuo琢,zhuo茁,zhuo酌,zhuo擢,zhuo焯,zhuo濯,zhuo诼,zhuo浞,zhuo涿,zhuo倬,zhuo镯,zhuo禚,zhuo斫,zhuo淖,zi字,zi自,zi子,zi紫,zi籽,zi资,zi姿,zi吱,zi滓,zi仔,"
		Pinyin_Dict = Pinyin_Dict & "zi兹,zi咨,zi孜,zi渍,zi滋,zi淄,zi笫,zi粢,zi龇,zi秭,zi恣,zi谘,zi趑,zi缁,zi梓,zi鲻,zi锱,zi孳,zi耔,zi觜,zi髭,zi赀,zi茈,zi訾,zi嵫,zi眦,zi姊,zi辎,zong总,zong纵,zong宗,zong棕,zong综,zong踪,zong鬃,zong偬,zong粽,zong枞,zong腙,zou走,"
		Pinyin_Dict = Pinyin_Dict & "zou揍,zou奏,zou邹,zou鲰,zou鄹,zou陬,zou驺,zou诹,zu组,zu族,zu足,zu阻,zu租,zu祖,zu诅,zu菹,zu镞,zu卒,zu俎,zuan钻,zuan纂,zuan缵,zuan躜,zuan攥,zui最,zui嘴,zui醉,zui罪,zui觜,zui蕞,zun尊,zun遵,zun鳟,zun撙,zun樽,zuo做,zuo作,zuo坐,zuo左,zuo座,"
		Pinyin_Dict = Pinyin_Dict & "zuo昨,zuo琢,zuo撮,zuo佐,zuo嘬,zuo酢,zuo唑,zuo祚,zuo胙,zuo怍,zuo阼,zuo柞,zuo砟,"


	'^^^^好长一段字典啊！！！

	RegExp_Match.IgnoreCase=True
	RegExp_Match.Global=True

	RegExp_PyChk.IgnoreCase=false
	RegExp_PyChk.Global=True

	For DictFinger=1 to Len(Pinyin_Dict)

		DictFinger_Char=Mid(Pinyin_Dict,DictFinger,1)
		If Asc(DictFinger_Char)<0 or Asc(DictFinger_Char)>255 Then
			RegExp_Match.Pattern="([a-z]+)[^a-z]*" & DictFinger_Char

			If RegExp_Match.Test(Pinyin_Dict)=True then
				set RegExp_MatchCollection = RegExp_Match.Execute(Pinyin_Dict)
				If RegExp_MatchCollection.count>=2 Then
					Pinyin=""
					Pinyin_Add_Count=0
					For I=0 to RegExp_MatchCollection.count-1
						Pinyin_Add = Ucase(Left(RegExp_MatchCollection.item(I).SubMatches.item(0),1)) & Mid(RegExp_MatchCollection.item(I).SubMatches.item(0),2,50) & ","
						RegExp_PyChk.Pattern=Pinyin_Add
						If RegExp_PyChk.Test(Pinyin)=False Then 
							Pinyin = Pinyin & Pinyin_Add
							Pinyin_Add_Count = Pinyin_Add_Count +1  
						End If
					next
					If Pinyin_Add_Count >=2 Then range("A" & DictFinger) = DictFinger_Char & ":" & Pinyin
				End If
			End If
		End IF
	Next

	Msgbox(Timer() - Time_ST)
End Sub