// (c) Ger Versluis 2000 version 4.16, 21 June 2003
//  You may use this script on non commercial sites
//  For info write to menu@burmees.nl

	var P_WW,P_WH;
	var P_RcrsLvl=0;
	var P_Ldd=0,	P_Crtd=0,	P_IniFlg;
	var P_FrstMnu=null,	P_CrrntOvr=null;
	var P_ClsTmr,	P_Ztp=100;
	var P_CntrTxt;
	var P_TxtCls;
	var P_show=Nav4?'show':'visible';
	var P_hide=Nav4?'hide':'hidden';
	var P_FStr="";
	P_WbMAlrts=["Item not defined: ","Item needs height: ","Item needs width: "];

	var AgntUsr=navigator.userAgent.toLowerCase();
	var AppVer=navigator.appVersion.toLowerCase();
	var DomYes=document.getElementById?1:0;
	var NavYes=AgntUsr.indexOf('mozilla')!=-1&&AgntUsr.indexOf('compatible')==-1?1:0;
	var ExpYes=AgntUsr.indexOf('msie')!=-1?1:0;
	var Opr=AgntUsr.indexOf('opera')!=-1?1:0;
	var DomNav=DomYes&&NavYes?1:0;
 	var DomExp=DomYes&&ExpYes?1:0;
	var Nav4=NavYes&&!DomYes&&document.layers?1:0;
	var Exp4=ExpYes&&!DomYes&&document.all?1:0;
	var P_Fltr=(AppVer.indexOf("msie 6")!= -1||AppVer.indexOf("msie 7")!= -1)?1:0;
	var PosStrt=(NavYes||ExpYes)&&!Opr?1:0;
	
	var P_Win,P_Doc,P_Bod;
	var P_ShwFlg=0;

function Pop_Go(){
	if(!PosStrt)return;
	P_BeforeStart();
	P_Win=window;
	P_Doc=P_Win.document;
	P_Bod=P_Doc.body;
	if(PopWebMasterCheck)
		if(!P_Check()){status='PopMenu build aborted';return}
	P_Crtd=0; P_Ldd=1;
	P_Create();
	P_Pos();
	P_Initiate();
	P_Win.onresize=Resize;
	if(ExpYes)P_Bod.onunload=P_FreeMem;
	P_Crtd=1;
	P_AfterBuild()}

function P_Check(){
	var WM='PopMenu',arr,i;	
	for(i=0;i<PopNoOffMenus;i++){
		arr=WM+eval(i+1);
		if(!P_Win[arr]){P_WAlrt(0,arr); return false}
		if(!P_ChckMn(arr+'_',P_Win[arr][0]))return false}
	return true}

function P_ChckMn(WMnu,NoOf){
	var i,Nof,arr;
	for(i=0;i<NoOf;i++){
		arr=WMnu+eval(i+1);
		if(!P_Win[arr]){P_WAlrt(0,arr); return false}
		Nof=P_Win[arr][3];
		if(i==0){	if(!P_Win[arr][4]){P_WAlrt(1,arr); return false}
			if(!P_Win[arr][5]){P_WAlrt(2,arr); return false}}
		if(Nof)if(!P_ChckMn(arr+'_',Nof))return false}
	return true}	

function P_WAlrt(No,Xtra){
	return confirm(P_WbMAlrts[No]+Xtra+'   ')}

function Resize(){
	if(Nav4&&(P_WW!=P_Win.innerWidth||P_WH!=P_Win.innerHeight))P_Win.location.reload();
	else P_Pos()}

function P_Pos(){
	P_WW=ExpYes?P_Bod.clientWidth:P_Win.innerWidth;
	P_WH=ExpYes?P_Bod.clientHeight:P_Win.innerHeight;
	var i,MPntr=P_FrstMnu,PreLft,PreTp,TP,Sz,PA;
	for(i=0;i<PopNoOffMenus;i++){
		PreLft=PreTp=0;
		PA=MPntr.PropArr;
		if(PA[20]){	if(DomYes){TP=P_Doc.getElementById(PA[20]);
				while(TP){PreTp+=TP.offsetTop;
					PreLft+=TP.offsetLeft;
					TP=TP.offsetParent}}
			else{	PreTp=Nav4?P_Doc.layers[PA[20]].pageY:P_Doc.all[PA[20]].offsetTop;
				PreLft=Nav4?P_Doc.layers[PA[20]].pageX:P_Doc.all[PA[20]].offsetLeft}}
		if(PA[22]!='left'){
			Sz=P_WW-(!Nav4?parseInt(MPntr.style.width):MPntr.clip.width);
			PreLft+=PA[22]=='right'?Sz:Sz/2}
		if(PA[23]!='top'){
			Sz=P_WH-(!Nav4?parseInt(MPntr.style.height):MPntr.clip.height);
			PreTp+=PA[23]=='bottom'?Sz:Sz/2}
		P_PosMenu(MPntr,(MPntr.StrtTp+PreTp),(MPntr.StrtLft+PreLft));
		MPntr=MPntr.PrvMnu}}

function P_PosMenu(CntPtr,Tp,Lt){
	var Tpi,Lefti,Hori, SubTp,SubLt,CCW;
	var Mmbr=CntPtr.FrstMmbr;
	var PA=CntPtr.PropArr,Bw=PA[14],Bbtw=PA[21],Hovl=PA[16],Vovl=PA[17];
	var P_PadLft=Mmbr.value.indexOf('<')==-1?DomNav?PA[40]:0:0;
	var P_PadTp=Mmbr.value.indexOf('<')==-1?DomNav?PA[39]:0:0;
	var MbrWdt=Nav4?Mmbr.clip.width:parseInt(Mmbr.style.width)+P_PadLft;
	var MbrHgt=Nav4?Mmbr.clip.height:parseInt(Mmbr.style.height)+P_PadTp;
	var CntWdt=Nav4?CntPtr.clip.width:parseInt(CntPtr.style.width);
	var CntHgt=Nav4?CntPtr.clip.height:parseInt(CntPtr.style.height);
	P_RcrsLvl++;
	if(Nav4)CntPtr.moveTo(Lt,Tp);
	else{CntPtr.style.top=Tp;CntPtr.style.left=Lt}
	CntPtr.OrgTp=Tp;CntPtr.OrgLft=Lt;
	if(P_RcrsLvl==1 && PA[12]){
		Hori=1;Lefti=CntWdt-MbrWdt-2*Bw;Tpi=0}
	else{	Hori=Lefti=0;Tpi=CntHgt-MbrHgt-2*Bw}
	while(Mmbr!=null){		
		if(Nav4){	Mmbr.moveTo(Lefti+Bw,Tpi+Bw);Mmbr.CmdLyr.moveTo(Lefti+Bw,Tpi+Bw)}
		else{	Mmbr.style.left=Lefti+Bw;Mmbr.style.top=Tpi+Bw}
		if(Mmbr.ChldCntnr){
			CCW=Nav4?Mmbr.ChldCntnr.clip.width:parseInt(Mmbr.ChldCntnr.style.width);
			if(Hori){	SubTp=Tpi+MbrHgt+Bw;SubLt=Lefti}
			else{	if(PA[19]){	SubLt=Lefti-CCW+Hovl*MbrWdt+Bw;SubTp=Tpi+(1-Vovl)*MbrHgt}
				else {	SubLt=Lefti+(1-Hovl)*MbrWdt+Bw;SubTp=Tpi+(1-Vovl)*MbrHgt}}
			P_PosMenu(Mmbr.ChldCntnr,SubTp,SubLt)}
		Mmbr=Mmbr.PrvMbr;
		if(Mmbr){	P_PadLft=Mmbr.value.indexOf('<')==-1?DomNav?PA[40]:0:0;
			P_PadTp=Mmbr.value.indexOf('<')==-1?DomNav?PA[39]:0:0;
			MbrWdt=Nav4?Mmbr.clip.width:parseInt(Mmbr.style.width)+P_PadLft;
			MbrHgt=Nav4?Mmbr.clip.height:parseInt(Mmbr.style.height)+P_PadTp;
			Hori?Lefti-=Bbtw?(MbrWdt+Bw):MbrWdt:Tpi-=Bbtw?(MbrHgt+Bw):MbrHgt}}
	P_RcrsLvl--}

function P_FreeMem(){
	var Mi, MPntr=P_FrstMnu;
	while(MPntr){
		Mi=MPntr;
		P_FreeMenu(MPntr);
		MPntr=MPntr.PrvMnu;
		Mi=null}
	P_FrstMnu=P_CrrntOvr=P_ClsTmr=null}

function P_FreeMenu(Cpntr){
	var Mi, Mbr=Cpntr.FrstMmbr;
	while(Mbr!=null){
		Mi=Mbr;
		if(Mbr.ChldCntnr) P_FreeMenu(Mbr.ChldCntnr);
		Mbr.ChldCntnr=null;
		Mbr.Contnr=null;
		Mbr=Mbr.PrvMbr;
		Mi.PrvMbr=null;
		Mi=null}}

function P_Initiate(){
	var MPntr=P_FrstMnu;
	var Mst;
	while(MPntr){
		P_ResetHide(MPntr);
		MPntr=MPntr.PrvMnu}}

function P_Reset(){
	if(!P_IniFlg)return;
	var ItemPntr=P_CrrntOvr.Contnr;
	while(ItemPntr.PrevCntnr) ItemPntr=ItemPntr.PrevCntnr;
	P_ResetHide(ItemPntr);
	if(P_ShwFlg)P_AfterCloseAll();P_ShwFlg=0}


function P_ResetHide(Cpntr){
	var Mbr=Cpntr.FrstMmbr;
	var Cst=Nav4?Cpntr:Cpntr.style;
	var PA=Cpntr.PropArr,bc;
	Cst.visibility=!(PA[13]&&Cpntr.Lvl==1)?P_hide:P_show;
	while(Mbr!=null){
		if(Mbr.hl){	Mbr.hl=0;
			if(PA[38]){
			bc=PA[Mbr.Lvl==1?4:25];
				if(Nav4){	if(Mbr.ro)Mbr.document.images[Mbr.rid].src=Mbr.ri1;
					else{	if(Mbr.value.indexOf('<img')==-1){
							if(bc)Mbr.bgColor=bc;
							Mbr.document.write(Mbr.value);
							Mbr.document.close()}}}
				else{	if(Mbr.ro)P_Doc.images[Mbr.rid].src=Mbr.ri1;
					else{	if(bc)Mbr.style.backgroundColor=bc;
						Mbr.style.color=PA[Mbr.Lvl==1?3:24]}}}}
		if(Mbr.ChldCntnr) P_ResetHide(Mbr.ChldCntnr);
		Mbr=Mbr.PrvMbr}}

function P_ClearAllChilds(Pntr){
	var CPstl,bc;
	var PA;
	while (Pntr){
		if(Pntr.hl){	Pntr.hl=0;
			PA=Pntr.Contnr.PropArr;
			if(PA[38]){
			bc=PA[Pntr.Lvl==1?4:25];
				if(Nav4){	if(Pntr.ro)Pntr.document.images[Pntr.rid].src=Pntr.ri1;
					else{	if(Pntr.value.indexOf('<img')==-1){
							if(bc)Pntr.bgColor=bc;
							Pntr.document.write(Pntr.value);
							Pntr.document.close()}}}
				else{	if(Pntr.ro)P_Doc.images[Pntr.rid].src=Pntr.ri1;
					else{	if(bc)Pntr.style.backgroundColor=bc;
						Pntr.style.color=PA[Pntr.Lvl==1?3:24]}}}
			if(Pntr.ChldCntnr){
				CPstl=Nav4?Pntr.ChldCntnr:Pntr.ChldCntnr.style;
				CPstl.visibility=P_hide;
				P_ClearAllChilds(Pntr.ChldCntnr.FrstMmbr)}
			break}
		Pntr=Pntr.PrvMbr}}	

function P_GoTo(){
	if(this.LinkTxt){
		var bc=this.Contnr.PropArr[this.Lvl==1?4:25];
		status=''; 
		if(Nav4){	if(bc)this.LowLyr.bgColor=bc;
			this.LowLyr.document.write(this.LowLyr.value);
			this.LowLyr.document.close()}
		else{	if(bc)this.style.backgroundColor=bc;
			this.style.color=this.Contnr.PropArr[this.Lvl==1?3:24]}
			this.LinkTxt.indexOf('javascript:')!=-1?eval(this.LinkTxt):P_Win.location.href=BaseHref+this.LinkTxt}}

function PopMenu(WMnu,Evnt){
	if(DomNav)Evnt.stopPropagation();
	if(!P_Ldd||!P_Crtd) return;
	var Tp,Lft,Pntr=null;
	var P_TpScrlld=ExpYes?P_Bod.scrollTop:P_Win.pageYOffset;
	var P_LftScrlld=ExpYes?P_Bod.scrollLeft:P_Win.pageXOffset;
	var EventX=Nav4?Evnt.pageX:Evnt.clientX+P_LftScrlld;
	var EventY=Nav4?Evnt.pageY:Evnt.clientY+P_TpScrlld;
	if(!Nav4){	WMnu+='_1';
		P_CrrntOvr=DomYes?P_Doc.getElementById(WMnu):P_Doc.all[WMnu];
		Pntr=DomYes?P_Doc.getElementById(WMnu+'c'):P_Doc.all[WMnu+'c']}
	else{	Pntr=P_FrstMnu;
		WMnu=PopNoOffMenus-WMnu.substr(7,WMnu.length-7);
		while(WMnu){Pntr=Pntr.PrvMnu;WMnu--}
		P_CrrntOvr=Pntr.FrstMmbr.CmdLyr}
	P_Initiate();
	var CntHt=Nav4?Pntr.clip.height:parseInt(Pntr.style.height);
	var CntWt=Nav4?Pntr.clip.width:parseInt(Pntr.style.width);
	var CntStl=Nav4?Pntr:Pntr.style;
	Tp=Pntr.OrgTp==-1?EventY:Pntr.OrgTp==-2?EventY-CntHt/2:Pntr.OrgTp;
	Lft=Pntr.OrgLft==-1?Pntr.PropArr[19]?EventX-CntWt:EventX:Pntr.OrgLft==-2?EventX-CntWt/2:Pntr.OrgLft;
	if((Pntr.OrgTp==-1||Pntr.OrgTp==-2)&&!Pntr.PropArr[13]){
		if(Tp+CntHt>P_WH+P_TpScrlld)Tp-=Pntr.OrgTp==-1?CntHt:CntHt/2;
		if(Lft+CntWt>P_WW+P_LftScrlld)Lft-=Pntr.OrgLft==-1?CntWt:CntWt/2;
		if(Tp<P_TpScrlld)Tp=P_TpScrlld;
		if(Lft<P_LftScrlld)Lft=P_LftScrlld}
	CntStl.top=Tp;
	CntStl.left=Lft;
	if(P_Fltr&&PopMenuSlide){Pntr.filters[0].Apply();Pntr.filters[0].play()}
	CntStl.visibility=P_show;
	P_IniFlg=0}

function P_OpenMenuClick(e){
	if(DomNav)e.stopPropagation();
	if(!P_Ldd||!P_Crtd) return;
	var PA=this.Contnr.PropArr,bc=PA[this.Lvl==1?6:27],x,y;
	if(P_CrrntOvr){x=P_CrrntOvr.Contnr; while(x.PrevCntnr)x=x.PrevCntnr;
		y=this.Contnr; while(y.PrevCntnr)y=y.PrevCntnr;
		x!=y&&x?P_ResetHide(x):P_ClearAllChilds(this.Contnr.FrstMmbr)}
	else P_ClearAllChilds(this.Contnr.FrstMmbr);
	P_CrrntOvr=this; P_IniFlg=0;
	if(Nav4){	this.LowLyr.hl=1;
		if(this.LowLyr.ro)this.LowLyr.document.images[this.LowLyr.rid].src=this.LowLyr.ri2;
		else{if(bc)this.LowLyr.bgColor=bc;
			if(this.LowLyr.value.indexOf('<img')==-1){
				this.LowLyr.document.write(this.LowLyr.Ovalue);
				this.LowLyr.document.close()}}}
	else{this.hl=1;if(this.ro)P_Win.document.images[this.rid].src=this.ri2;
		else{if(bc)this.style.backgroundColor=bc;this.style.color=PA[this.Lvl==1?5:26]}}
	status=this.LinkTxt}	

function P_OpenMenu(e){
	if(DomNav)e.stopPropagation();
	if(!P_Ldd||!P_Crtd) return;
	var PA=this.Contnr.PropArr;
	var bc=PA[this.Lvl==1?6:27];
	var Lft,Tp,x,y;
	var P_TpScrlld=ExpYes?P_Bod.scrollTop:P_Win.pageYOffset;
	var P_LftScrlld=ExpYes?P_Bod.scrollLeft:P_Win.pageXOffset;
	var ChldCont=Nav4?this.LowLyr.ChldCntnr:this.ChldCntnr;
	var ContTp=Nav4?this.Contnr.top:parseInt(this.Contnr.style.top);
	var ContLft=Nav4?this.Contnr.left:parseInt(this.Contnr.style.left);
	var CntWt=Nav4?this.Contnr.clip.width:parseInt(this.Contnr.style.width);
	var ThisHt=Nav4?this.clip.height:parseInt(this.style.height);
	var ThisWt=Nav4?this.clip.width:parseInt(this.style.width);
	if(P_CrrntOvr){
		x=P_CrrntOvr.Contnr; while(x.PrevCntnr)x=x.PrevCntnr;
		y=this.Contnr; while(y.PrevCntnr)y=y.PrevCntnr;
		x!=y&&x?P_ResetHide(x):P_ClearAllChilds(this.Contnr.FrstMmbr)}
	else P_ClearAllChilds(this.Contnr.FrstMmbr);
	P_CrrntOvr=this; P_IniFlg=0;
	if(Nav4){	this.LowLyr.hl=1;
		if(this.LowLyr.ro)this.LowLyr.document.images[this.LowLyr.rid].src=this.LowLyr.ri2;
		else{	if(bc)this.LowLyr.bgColor=bc;
			if(this.LowLyr.value.indexOf('<img')==-1){
				this.LowLyr.document.write(this.LowLyr.Ovalue);
				this.LowLyr.document.close()}}}
	else{	this.hl=1;
		if(this.ro)P_Win.document.images[this.rid].src=this.ri2;
		else{
			if(bc)this.style.backgroundColor=bc;
			this.style.color=PA[this.Lvl==1?5:26]}}
	if(ChldCont!=null){
		if(!P_ShwFlg){P_ShwFlg=1;P_BeforeFirstOpen()}
		var CCW=Nav4?this.LowLyr.ChldCntnr.clip.width:parseInt(this.ChldCntnr.style.width);
		var CCH=Nav4?this.LowLyr.ChldCntnr.clip.height:parseInt(this.ChldCntnr.style.height);
		var CCSt=Nav4?this.LowLyr.ChldCntnr:this.ChldCntnr.style;
		Tp=ChldCont.OrgTp+ContTp;
		Lft=ChldCont.OrgLft+ContLft;
		if(PA[19]){
			if(Lft<P_LftScrlld)Lft=PA[12]&&this.Contnr.Lvl==1?P_LftScrlld:Lft+(CCW+(1-2*PA[16])*ThisWt);
			if(Lft+CCW>P_WW+P_LftScrlld)Lft=P_WW+P_LftScrlld-CCW}
		else{	if(Lft+CCW>P_WW+P_LftScrlld)Lft=PA[12]&&this.Contnr.Lvl==1?P_WW+P_LftScrlld-CCW:Lft-(CCW+(1-2*PA[16])*ThisWt);
			if(Lft<P_LftScrlld)Lft=P_LftScrlld}
		if(Tp+CCH>P_WH+P_TpScrlld)Tp=Tp-CCH-(1-2*PA[17])*ThisHt;
		if(Tp<P_TpScrlld)Tp=P_TpScrlld;
		CCSt.left=Lft;
		CCSt.top=Tp;
		if(P_Fltr&&PopMenuSlide){this.ChldCntnr.filters[0].Apply();this.ChldCntnr.filters[0].play()}
		CCSt.visibility=P_show}
	status=this.LinkTxt}	

function OutMenu(WMnu){
	if(!P_Ldd||!P_Crtd||!P_CrrntOvr)return;
	P_IniFlg=1;
	if (P_ClsTmr) clearTimeout(P_ClsTmr);
	P_ClsTmr=setTimeout('P_Reset()',P_Win[WMnu][18])}

function P_CloseMenu(e){
	if(DomNav)e.stopPropagation();
	if(!P_Ldd||!P_Crtd) return;
	var PA=this.Contnr.PropArr;
	var bc=PA[this.Lvl==1?4:25];
	if(!PA[38]){
		if(Nav4){	if(this.LowLyr.ro)this.LowLyr.document.images[this.LowLyr.rid].src=this.LowLyr.ri1;
			else{	if(this.LowLyr.value.indexOf('<img')==-1){
					if(bc)this.LowLyr.bgColor=bc;
					this.LowLyr.document.write(this.LowLyr.value);
					this.LowLyr.document.close()}}}
		else{	if(this.ro)P_Win.document.images[this.rid].src=this.ri1;
			else{	if(bc)this.style.backgroundColor=bc;
				this.style.color=PA[this.Lvl==1?3:24]}}}
	status='';
	P_IniFlg=1;
	if (P_ClsTmr) clearTimeout(P_ClsTmr);
	P_ClsTmr=setTimeout('P_Reset()',PA[18])}

function P_CntnrSetUp(Wdth,Hght,NoOff,Lft,Tp,PCntnr){
	var PA=this.PropArr;
	var bc=P_RcrsLvl==1?7:28;
	if(Nav4){	this.visibility='hide';
		this.zIndex=P_RcrsLvl+P_Ztp}
	this.FrstMmbr=null;
	this.PrvMnu=null;
	this.PrevCntnr=PCntnr;
	this.StrtLft=this.OrgLft=Lft;
	this.StrtTp=this.OrgTp=Tp;
	this.Lvl=P_RcrsLvl;
	if(PA[bc]){
		if(Nav4)this.bgColor=PA[bc];
		else this.style.backgroundColor=PA[bc]}
	if(!Nav4){	this.style.width=Wdth;
		this.style.height=Hght}
	else this.resizeTo(Wdth,Hght);
	if(!Nav4){	with(this.style){
			fontFamily=PA[8];
			fontWeight=PA[9]?'bold':'normal';
			fontStyle=PA[10]?'italic':'normal';
			fontSize=PA[11]+'pt';
			zIndex=P_RcrsLvl+P_Ztp;
			top=-1000;
			left=-1000}}
	if(P_Fltr){P_FStr="";if(PopMenuSlide&&!(P_RcrsLvl==1&&PA[13]))P_FStr=PopMenuSlide;if(PopMenuShadow)P_FStr+=PopMenuShadow;
	if(PopMenuOpacity)P_FStr+=PopMenuOpacity;if(P_FStr!="")this.style.filter=P_FStr}}

function P_MemberSetUp(MmbrCntnr,PrMmbr,WMnu,Wdth,Hght){
	var MemVal=eval(WMnu+'[0]');
	var t,T,L,W,H,S;
	var PA=MmbrCntnr.PropArr;
	var tri=P_RcrsLvl==1&&PA[12]?32:PA[19]?35:29;
	this.ro=0;
	if(MemVal.indexOf('rollover')!=-1){
		this.ro=1;
		this.ri1=MemVal.substring(MemVal.indexOf('?')+1,MemVal.lastIndexOf('?'));
		this.ri2=MemVal.substring(MemVal.lastIndexOf('?')+1,MemVal.length);
		this.rid=WMnu+'i';
		MemVal="<img src='"+this.ri1+"' name='"+this.rid+"'>"}
	this.value=MemVal;
	this.ChldCntnr=null;
	this.PrvMbr=PrMmbr;
	this.LinkTxt=eval(WMnu+'[1]');
	this.Lvl=P_RcrsLvl;
	this.hl=0;
	with(this.style){
		if(MemVal.indexOf('<')==-1){
			width=Wdth-(DomNav?PA[40]:0);
			height=Hght-(DomNav?PA[39]:0);
			paddingLeft=PA[40];
			paddingTop=PA[39]}
		else{	width=Wdth;
			height=Hght}
		overflow='hidden';
		cursor=this.LinkTxt?ExpYes?"hand":"pointer":"default";
		if(PA[P_RcrsLvl==1?4:25])backgroundColor=PA[P_RcrsLvl==1?4:25];
		color=PA[this.Lvl==1?3:24];
		if(PA[15]!='left')textAlign=PA[15]}
	if(eval(WMnu+'[2]'))this.style.backgroundImage="url(\""+eval(WMnu+'[2]')+"\")";
	if(MemVal.indexOf('<')==-1&&DomYes){var t=P_Doc.createTextNode(MemVal);this.appendChild(t)}
	else this.innerHTML=MemVal;
	if(eval(WMnu+'[3]')){
		S=PA[tri];
		W=PA[tri+1];
		H=PA[tri+2];
		T=P_RcrsLvl==1&&PA[12]?Hght-H-2:(Hght-H)/2;
		L=PA[19]?2:Wdth-W-2;
		if(DomYes){
			t=P_Doc.createElement('img');
			this.appendChild(t);
			t.style.position='absolute';
			t.src=S;
			t.style.width=W;
			t.style.height=H;
			t.style.top=T;
			t.style.left=L}
		else{	MemVal+="<div style='position:absolute; top:"+T+"; left:"+L+"; width:"+W+"; height:"+H+";visibility:inherit'><img src='"+S+"'></div>";
			this.innerHTML=MemVal}}
	if(DomNav){PA[41]&&P_RcrsLvl==1?this.addEventListener('mouseover',P_OpenMenuClick,false):this.addEventListener('mouseover',P_OpenMenu,false);
		this.addEventListener('mouseout',P_CloseMenu,false);
		PA[41]&&P_RcrsLvl==1?this.addEventListener('click',P_OpenMenu,false):this.addEventListener('click',P_GoTo,false)}
	else{	this.onmouseover=PA[41]&&P_RcrsLvl==1?P_OpenMenuClick:P_OpenMenu;
		this.onmouseout=P_CloseMenu;
		this.onclick=PA[41]&&P_RcrsLvl==1?P_OpenMenu:P_GoTo}
	this.Contnr=MmbrCntnr}

function P_Nav_MemberSetUp(MmbrCntnr,PrMmbr,WMnu,Wdth,Hght){
	var PA=MmbrCntnr.PropArr;
	var tri=P_RcrsLvl==1&&PA[12]?32:PA[19]?35:29;
	this.value=eval(WMnu+'[0]');
	this.ro=0;
	if(this.value.indexOf('rollover')!=-1){
		this.ro=1;
		this.ri1=this.value.substring(this.value.indexOf('?')+1,this.value.lastIndexOf('?'));
		this.ri2=this.value.substring(this.value.lastIndexOf('?')+1,this.value.length);
		this.rid=WMnu+'i';this.value="<img src='"+this.ri1+"' name='"+this.rid+"'>"}
	if(PA[40]&&this.value.indexOf('<')==-1&&PA[15]=='left')this.value='&nbsp\;'+this.value;
	if(PA[9])this.value=this.value.bold();
	if(PA[10])this.value=this.value.italics();
	this.Ovalue=this.value;
	this.value=this.value.fontcolor(PA[P_RcrsLvl==1?3:24]);
	this.Ovalue=this.Ovalue.fontcolor(PA[P_RcrsLvl==1?5:26]);
	this.value=P_CntrTxt+"<font face='"+PA[8]+"' point-size='"+PA[11]+"'>"+this.value+P_TxtCls;
	this.Ovalue=P_CntrTxt+"<font face='"+PA[8]+"' point-size='"+PA[11]+"'>"+this.Ovalue+P_TxtCls;
	this.visibility='inherit';
	this.ChldCntnr=null;
	this.PrvMbr=PrMmbr;
	this.Contnr=MmbrCntnr;
	this.Lvl=P_RcrsLvl;
	this.hl=0;
	if(PA[P_RcrsLvl==1?4:25])this.bgColor=PA[P_RcrsLvl==1?4:25];
	this.resizeTo(Wdth,Hght);
	if(eval(WMnu+'[2]'))this.background.src=eval(WMnu+'[2]');
	this.document.write(this.value);
	this.document.close();
	this.CmdLyr=new Layer(Wdth,MmbrCntnr);
	this.CmdLyr.visibility='inherit';
	this.CmdLyr.Lvl=P_RcrsLvl;
	this.CmdLyr.LinkTxt=eval(WMnu+'[1]');
	this.CmdLyr.onmouseover=PA[41]&&P_RcrsLvl==1?P_OpenMenuClick:P_OpenMenu;
	this.CmdLyr.onmouseout=P_CloseMenu;
	this.CmdLyr.captureEvents(Event.MOUSEUP);
	this.CmdLyr.onmouseup=PA[41]&&P_RcrsLvl==1?P_OpenMenu:P_GoTo;
	this.CmdLyr.LowLyr=this;
	this.CmdLyr.Contnr=MmbrCntnr;
	this.CmdLyr.resizeTo(Wdth,Hght);
	if(eval(WMnu+'[3]')){
		this.CmdLyr.ImgLyr=new Layer(PA[tri+1],this.CmdLyr);
		this.CmdLyr.ImgLyr.visibility='inherit';
		this.CmdLyr.ImgLyr.top=P_RcrsLvl==1&&PA[12]?Hght-PA[tri+2]-2:(Hght-PA[tri+2])/2;
		this.CmdLyr.ImgLyr.left=PA[19]?2:Wdth-PA[tri+1]-2;
		this.CmdLyr.ImgLyr.width=PA[tri+1];
		this.CmdLyr.ImgLyr.height=PA[tri+2];
		ImgStr="<img src='"+PA[tri]+"' width='"+PA[tri+1]+"' height='"+PA[tri+2]+"'>";
		this.CmdLyr.ImgLyr.document.write(ImgStr);
		this.CmdLyr.ImgLyr.document.close()}}

function P_Create(){
	var i,algn;
	var WMnu,MPntr,MenuPrevPntr=null;
	for(i=0;i<PopNoOffMenus;i++){
		WMnu='PopMenu'+(i+1);
		algn=eval(WMnu+'[15]');
		P_CntrTxt=DomYes?algn:Exp4?algn!='left'?"align='"+algn+"'":"":algn!='left'?"<div align='"+algn+"'>":"";
		if(Nav4)P_TxtCls="</font>"+(algn!='left'?"</div>":"");
		MPntr=P_CreateMenuStructure(WMnu,WMnu+'_',eval(WMnu+'[0]'),eval(WMnu+'[1]'),eval(WMnu+'[2]'),null);
		MPntr.PrvMnu=MenuPrevPntr;
		MenuPrevPntr=MPntr}
	P_FrstMnu=MPntr}

function P_CreateMenuStructure(ArrPntr,MName,NmbOf,Lft,Tp,PrvCntnr){
	P_RcrsLvl++;
	var i,NSubs,Mmbr,MmbrCntnr,Wdth=0,Hght=0;
	var PrvMmbr=null;
	var WMnu=MName+'1';
	var MnWdth=eval(WMnu+'[5]');
	var MnHght=eval(WMnu+'[4]');
	var InsertLoc;
	var AP=P_Win[ArrPntr];
	if(P_RcrsLvl==1&&AP[12]){
		for(i=1;i<NmbOf+1;i++){
			WMnu=MName+eval(i);
			Wdth=eval(WMnu+'[5]')?Wdth+eval(WMnu+'[5]'):Wdth+MnWdth}
		Wdth=AP[21]?Wdth+(NmbOf+1)*AP[14]:Wdth+2*AP[14];Hght=MnHght+2*AP[14]}
	else{	for(i=1;i<NmbOf+1;i++){
			WMnu=MName+eval(i);
			Hght=eval(WMnu+'[4]')?Hght+eval(WMnu+'[4]'):Hght+MnHght}
		Hght=AP[21]?Hght+(NmbOf+1)*AP[14]:Hght+2*AP[14];Wdth=MnWdth+2*AP[14]}
	WMnu=MName+'1';	
	if(!Nav4)WMnu+='c';
	if(DomYes){
		MmbrCntnr=P_Doc.createElement("div");
		MmbrCntnr.style.visibility='hidden';
		MmbrCntnr.id=WMnu;
		MmbrCntnr.style.position='absolute';
		P_Bod.appendChild(MmbrCntnr)}
	else        if(Nav4)	MmbrCntnr=new Layer(Wdth,P_Win);
		else{	P_Bod.insertAdjacentHTML("AfterBegin","<div id='"+WMnu+"' style='visibility:hidden; position:absolute'><\/div>");
			MmbrCntnr=P_Doc.all[WMnu]}
	MmbrCntnr.SetUp=P_CntnrSetUp;
	MmbrCntnr.PropArr=AP;
	MmbrCntnr.SetUp(Wdth,Hght,NmbOf,Lft,Tp,PrvCntnr);
	if(Exp4){	MmbrCntnr.InnerString='';
		for(i=1;i<NmbOf+1;i++){
			WMnu=MName+eval(i);
			NSubs=eval(WMnu+'[3]');
			MmbrCntnr.InnerString+="<div id='"+WMnu+"' style='position:absolute;'><\/div>"}
		MmbrCntnr.innerHTML=MmbrCntnr.InnerString}
	for(i=1;i<NmbOf+1;i++){
		WMnu=MName+eval(i);
		NSubs=eval(WMnu+'[3]');
	Wdth=P_RcrsLvl==1&&AP[12]?eval(WMnu+'[5]')?eval(WMnu+'[5]'):MnWdth:MnWdth;
	Hght=P_RcrsLvl==1&&AP[12]?MnHght:eval(WMnu+'[4]')?eval(WMnu+'[4]'):MnHght;
	if(DomYes){Mmbr=P_Doc.createElement("div");
		Mmbr.style.position='absolute';
		Mmbr.style.visibility='inherit';
		Mmbr.id=WMnu;
		MmbrCntnr.appendChild(Mmbr);
		Mmbr.SetUp=P_MemberSetUp}
	else 	if(Nav4){	Mmbr=new Layer(Wdth,MmbrCntnr);
			Mmbr.SetUp=P_Nav_MemberSetUp}
		else{	Mmbr=MmbrCntnr.all[WMnu];
			Mmbr.SetUp=P_MemberSetUp}
		Mmbr.SetUp(MmbrCntnr,PrvMmbr,WMnu,Wdth,Hght);
		if(NSubs) Mmbr.ChldCntnr=P_CreateMenuStructure(ArrPntr,WMnu+'_',NSubs,0,0,MmbrCntnr);
		PrvMmbr=Mmbr}
	MmbrCntnr.FrstMmbr=Mmbr;
	P_RcrsLvl--;
	return(MmbrCntnr)}