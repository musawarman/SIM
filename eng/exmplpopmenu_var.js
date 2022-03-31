//(c) Ger Versluis 2000 version 4.15, 9 July 2002
// Notation of PopMenu1 is different from PopMenu2. The result is the same. PopMenu1 is more understandable. PopMenu2 loads faster.

	// Globals
	var PopNoOffMenus=3;
	var PopWebMasterCheck=0;
	var BaseHref="";

	var PopMenuSlide="";
	var PopMenuSlide="progid:DXImageTransform.Microsoft.RevealTrans(duration=.25, transition=18)";
	var PopMenuSlide="progid:DXImageTransform.Microsoft.GradientWipe(duration=.25, wipeStyle=1)";

	var PopMenuShadow="";
	//var PopMenuShadow="progid:DXImageTransform.Microsoft.DropShadow(color=#888888, offX=2, offY=2, positive=1)";
	//var PopMenuShadow="progid:DXImageTransform.Microsoft.Shadow(color=#888888, direction=135, strength=0)";

	var PopMenuOpacity="";
	var PopMenuOpacity="progid:DXImageTransform.Microsoft.Alpha(opacity=90)";

	
	function P_BeforeStart(){return}
	function P_AfterBuild(){return}
	function P_BeforeFirstOpen(){return}
	function P_AfterCloseAll(){return}

PopMenu1=new Array(2,201,133,"000000","F4BC31","F4BC31","000000","000000","arial",0,0,8,0,0,1,"left",-0.01,1,1000,0,"",1,"left","top","000000","F4BC31","F4BC31","000000","000000",BaseHref+"tri.gif",5,10,BaseHref+"tridown.gif",10,5,BaseHref+"trileft.gif",5,10,1,2,2,0);	
	PopMenu1_1=new Array("Poultry Division","business/index.html","",5,24,137);
		PopMenu1_1_1=new Array("Feedmill","business/feedmill.html","",0,24,137);
		PopMenu1_1_2=new Array("Breeding","business/breeding.html","",0,24,137);
		PopMenu1_1_3=new Array("Slaughterhouse","business/slaught.html","",0,24,137);
		PopMenu1_1_4=new Array("Equipment","business/equipment.html","",0,24,137);
		PopMenu1_1_5=new Array("Animal Pharmaceutical","business/animal.html","",0,24,137);
	PopMenu1_2=new Array("Franchise Restaurant","business/franchise.html","",2,24,137);
		PopMenu1_2_1=new Array("Wendy's","business/franchise1.html","",0,24,137);
		PopMenu1_2_2=new Array("Hartz","business/franchise1.html","",0,24,137);

PopMenu2=new Array(2,201,103,"000000","F4BC31","F4BC31","000000","000000","arial",0,0,8,0,0,1,"left",-0.01,1,1000,0,"",1,"left","top","000000","F4BC31","F4BC31","000000","000000",BaseHref+"tri.gif",5,10,BaseHref+"tridown.gif",10,5,BaseHref+"trileft.gif",5,10,1,2,2,0);	
	PopMenu2_1=new Array("Vision & Mision","corporate/index.html","",0,24,137);
	PopMenu2_2=new Array("History Background","corporate/history.html","",0,24,137);

PopMenu3=new Array(6,201,165,"000000","F4BC31","F4BC31","000000","000000","arial",0,0,8,0,0,1,"left",-0.01,1,1000,0,"",1,"left","top","000000","F4BC31","F4BC31","000000","000000",BaseHref+"tri.gif",5,10,BaseHref+"tridown.gif",10,5,BaseHref+"trileft.gif",5,10,1,2,2,0);	
	PopMenu3_1=new Array("Feed","products/feed.html","",0,24,137);
	PopMenu3_2=new Array("DOC","products/doc.html","",0,24,137);
	PopMenu3_3=new Array("Equipment","products/equipment.html","",0,24,137);
	PopMenu3_4=new Array("Animal Pharmaceutical","products/animal.html","",0,24,137);
	PopMenu3_5=new Array("Dressed Chickent","products/dressed.html","",0,24,137);
	PopMenu3_6=new Array("Delfarm","products/delfarm.html","",0,24,137);
