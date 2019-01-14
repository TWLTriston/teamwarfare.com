function getScreenWidth() {
	if (!(bw)) {
		bw = new cm_bwcheck();
	}
	if (bw.ie) {
		return document.body.offsetWidth - 20;
	} else if (bw.moz) {
		return document.body.offsetWidth - 10;
	} else if (bw.op7) {
		return document.body.offsetWidth;
	} else if (bw.ns4 || bw.ns6) {
		return window.innerWidth;
	} else {
		return 0;
	}
}


function cm_bwcheck(){
	this.ver=navigator.appVersion
	this.agent=navigator.userAgent.toLowerCase()
	this.dom=document.getElementById?1:0
	this.op5=(this.agent.indexOf("opera 5")>-1 || this.agent.indexOf("opera/5")>-1) && window.opera 
	this.op6=(this.agent.indexOf("opera 6")>-1 || this.agent.indexOf("opera/6")>-1) && window.opera   
	this.op7=(this.agent.indexOf("opera 7")>-1 || this.agent.indexOf("opera/7")>-1) && window.opera   
	this.moz=(this.agent.indexOf("mozilla")>-1 );
	this.ie5 = (this.agent.indexOf("msie 5")>-1 && !this.op5 && !this.op6)
	this.ie55 = (this.ie5 && this.agent.indexOf("msie 5.5")>-1)
	this.ie6 = (this.agent.indexOf("msie 6")>-1 && !this.op5 && !this.op6)
	this.ie4=(this.agent.indexOf("msie")>-1 && document.all &&!this.op5 &&!this.op6 &&!this.ie5&&!this.ie6)
	this.ie = (this.ie4 || this.ie5 || this.ie6)
	this.mac=(this.agent.indexOf("mac")>-1)
	this.ns6=(this.agent.indexOf("gecko")>-1 || window.sidebar)
	this.ns4=(!this.dom && document.layers)?1:0;
	this.bw=(this.ie6 || this.ie5 || this.ie4 || this.ns4 || this.ns6 || this.op5 || this.op6)
	this.usedom= this.ns6//Use dom creation
	this.reuse = this.ie||this.usedom //Reuse layers
	this.px=this.dom&&!this.op5?"px":""
	return this
}

var bw = new cm_bwcheck();

function fCenterMenus() {
	intDocumentWidth = 990
	iWinWidth = getScreenWidth();
	if (iWinWidth >= intDocumentWidth) {
		iLeftCoor = (iWinWidth / 2) - (intDocumentWidth / 2);
	} else {
		iLeftCoor = 0;
	}
	document.getElementById("divContainer").style.left = iLeftCoor;
	document.getElementById("divContainer").style.visibility = "visible";
	document.getElementById("divMenu").style.visibility = "visible";
	document.getElementById("divMenumytwl").style.visibility = "visible";
	document.getElementById("divMenuforums").style.visibility = "visible";
	document.getElementById("divMenuoperations").style.visibility = "visible";
	document.getElementById("divMenuhelp").style.visibility = "visible";
}

var mi = new Array();

var arrParents = new Array();
var arrMyTWL = new Array();
var arrForums = new Array();
var arrComp = new Array();
var arrRules = new Array();
var arrOperations = new Array();
var arrHelp = new Array();
var arrStaffForums = new Array();
var arrStaffForums2 = new Array();

function clsStyle() {
	this.DisplayWide = false; // vs true
	this.OnBackgroundColor = "#550000";
	this.OnColor = "#ffd142";
	this.OffBackgroundColor = "#3C0000";
	this.OffColor = "#ffffff";
	this.Width = 200;
	this.Height = 23;
	this.LeftOffset = -105;
	this.TopOffset = 0;
	this.ShowArrow = true;
	this.ArrowLeft = 186;
	this.ArrowTop = 7;
}
	
function clsMenuItem() {
	this.Name = "";
	this.dn = "";
	this.url = "";
	this.jurl = "";
	this.Children = 0;
	this.PName = "";
	this.PArrIndex = 0;
	this.PLeft = 0;
	this.PTop = 0;
	this.PTier = 0;
	this.Left = 0;
	this.Top = 0;
	this.t = 0;
	this.IN = -1;
}

function clsMenu() {
	this.Name = "";
	this.Top = 0;
	this.Left = 0;
	this.t = 0;
}
var iAlerted = 0;
function fMakeMenu(arrParent, strName, strPName, strDisplay, strLink, strJavascriptURL) {
	
	var iIndex = arrParent.push(new clsMenuItem()) - 1;
	arrParent[iIndex].Name = strName;
	arrParent[iIndex].PName = strPName;
	arrParent[iIndex].dn = strDisplay;
	arrParent[iIndex].jurl = strJavascriptURL;
	arrParent[iIndex].url = strLink;
	for (i=0;i<arrParent.length;i++) {
		if (arrParent[i].PName == strPName) {
			arrParent[iIndex].IN++;
		}
	}
	if (strPName != "") {
		for (i=0;i<arrParent.length;i++) {
			if (arrParent[i].Name == strPName) {
				arrParent[i].Children++;
				arrParent[iIndex].PArrIndex = i;
				arrParent[iIndex].PTier = arrParent[i].t;
				arrParent[iIndex].t = arrParent[i].t + 1;
				arrParent[iIndex].PLeft = arrParent[i].Left + arrParent[i].PLeft + arrStyles[arrParent[iIndex].t].LeftOffset;
				arrParent[iIndex].PTop = arrParent[i].Top + arrParent[i].PTop + arrStyles[arrParent[iIndex].t].TopOffset;
			}
		}
	}

	if (arrStyles[arrParent[iIndex].t].DisplayWide) {
		arrParent[iIndex].Left = arrParent[iIndex].IN * arrStyles[arrParent[iIndex].t].Width;
		arrParent[iIndex].Top = 0;
	} else {
		arrParent[iIndex].Top = arrParent[iIndex].IN * arrStyles[arrParent[iIndex].t].Height;
		arrParent[iIndex].Left = 0;
	}
	if (arrParent[iIndex].t > 0) {
		if (arrStyles[arrParent[iIndex].PTier].DisplayWide) {
			arrParent[iIndex].Top = arrParent[iIndex].Top + arrStyles[arrParent[iIndex].PTier].Height;
		} else {
			arrParent[iIndex].Left = arrParent[iIndex].Left + arrStyles[arrParent[iIndex].PTier].Width;
		}
	}

}

function fInitMenus(arrParent, m, iLeft) {
	m.push (new clsMenu());
	m[0].Children = 6;
	for (i=0;i<arrParent.length;i++) {
		if (arrParent[i].Children > 0) {
			j = m.push(new clsMenu()) - 1;

			m[j].t = arrParent[i].t+1;
			m[j].Name = arrParent[i].Name;
			m[j].Top = arrParent[i].Top + arrParent[i].PTop; 
			m[j].Left = arrParent[i].Left + arrParent[i].PLeft + iLeft;
		} 
	}
}

function fDrawMenus(arrParent, strName, iLeft) {
	var m = new Array();
	fInitMenus(arrParent, m, iLeft);
	m[0].Name = m[0].Name + strName;
	iAlerted = 0;
	for (j=0;j<m.length;j++) {
		var arrArrowTop = new Array();
		var arrArrowLeft = new Array();
		var strOnClick = "";
		strOutput = '<div id="divMenu' + m[j].Name + '" style="position:absolute;top:' + m[j].Top + ';left:' + m[j].Left + ';visibility:hidden;">';
		for (i=0;i<arrParent.length;i++) {
			if (arrParent[i].PName == m[j].Name) {
				strOnClick = "";
				if (arrParent[i].url) {
					strOnClick = ' onClick="window.location=\'/' + arrParent[i].url + '\';"';
				} else if (arrParent[i].jurl) {
					strOnClick = ' onClick="' + arrParent[i].jurl + '"';
				}
				strOutput = strOutput + '<div style="position:absolute;top:' + arrParent[i].Top + 'px;left:' + arrParent[i].Left + 'px;" id="divMenuItem' + i + '" class="cssTier' + arrParent[i].t + 'Off" onMouseOut="fMouseOut();" onMouseOver="fMouseOver(this, \'divMenu' + arrParent[i].Name + '\', ' + m[j].t + ');"' + strOnClick + '>' + arrParent[i].dn + '</div>';
				if ((arrParent[i].Children > 0) && (arrStyles[arrParent[i].t].ShowArrow)){
					arrArrowTop.push(arrParent[i].Top + arrStyles[arrParent[i].t].ArrowTop);
					arrArrowLeft.push(arrParent[i].Left + arrStyles[arrParent[i].t].ArrowLeft); 
				}
			}
		}
		for (p=0;p<arrArrowLeft.length;p++) {
			strOutput = strOutput + '<div style="position:absolute;top:' + (arrArrowTop[p]) + ';left:' + (arrArrowLeft[p])+ ';visibility:inherit;z-order: 100;">';
			strOutput = strOutput + '<img src="/images/tri.gif" height="9" width="10" alt="" border="0" />';
			strOutput = strOutput + '</div>';
		}
		strOutput = strOutput + '</div>';
		document.write (strOutput);
	}
}

var arrDivs = new Array();
var oTimer = null;

function fMouseOver(oDiv, strMenuName, iTier) {
	// Highlight our selection
	window.clearTimeout(oTimer);
	fHide(iTier);
	arrDivs.push(new Array());
	arrDivs[arrDivs.length - 1][0] = oDiv
	arrDivs[arrDivs.length - 1][1] = iTier
	arrDivs[arrDivs.length - 1][2] = strMenuName
	
	oDiv.style.background = arrStyles[iTier].OnBackgroundColor;
	oDiv.style.color = arrStyles[iTier].OnColor;
	// Show any children
	if (document.getElementById(strMenuName)) {
		document.getElementById(strMenuName).style.visibility = "visible";
	}
	if (document.getElementById("divFlash")) {
		document.getElementById("divFlash").style.visibility = "hidden";
	}
}

function fHide(iTier) {
	if (arrDivs.length > 0) {
		while (arrDivs.length > 0) {
			// Find everything that is at this tier level, and turn it off
			if (arrDivs[arrDivs.length - 1][1] == iTier) {
				var arrThisDiv = arrDivs.pop();
				arrThisDiv[0].style.background = arrStyles[arrThisDiv[1]].OffBackgroundColor;
				arrThisDiv[0].style.color = arrStyles[arrThisDiv[1]].OffColor;
				strMenuName = arrThisDiv[2];
				if (document.getElementById(strMenuName)) {
					document.getElementById(strMenuName).style.visibility = "hidden";
				}
			} else if (arrDivs[arrDivs.length - 1][1] > iTier) {
				var arrThisDiv = arrDivs.pop();
				arrThisDiv[0].style.background = arrStyles[arrThisDiv[1]].OffBackgroundColor;
				arrThisDiv[0].style.color = arrStyles[arrThisDiv[1]].OffColor;
				strMenuName = arrThisDiv[2];
				if (document.getElementById(strMenuName)) {
					document.getElementById(strMenuName).style.visibility = "hidden";
				}
			} else { break }
		}
	}
	if (iTier == 0) {
		if (document.getElementById("divFlash")) {
			document.getElementById("divFlash").style.visibility = "visible";
		}
	}
}

function fMouseOut() {
	oTimer = window.setTimeout("fHide(0);", 300);
}

var arrStyles = new Array(new clsStyle(), new clsStyle(), new clsStyle(), new clsStyle());

arrStyles[0].DisplayWide = true;	
arrStyles[0].Width = 165;
arrStyles[0].ShowArrow = false;
arrStyles[0].Height = 24;

arrStyles[2].Height = 21;
arrStyles[3].Height = 21;

// OLD JS
var tri = new Image();
function preload() {
	tri.src = "/images/tri.gif"
}

function popup(url, name, height, width, scrollbars)
{
	var popwin;
	var opts = "toolbar=no,status=no,location=no,menubar=no,resizable=no";
	opts += ",height=" + height + ",width=" + width + ",scrollbars=" + scrollbars;
	popwin = window.open(url, name, opts);
	popwin.focus();
}

function fPopLogin() {
	var oLoginWindow;
	oLoginWindow = window.open("/login.asp?url=" + escape(String(window.location)), "LoginWindow", "height=175,width=300,scrollbars=0,resizable=0,menubar=0,location=0,status=0,toolbar=0");
	oLoginWindow.focus();
}

function fPopHelp() {
	var oHelpWindow;
	oHelpWindow = window.open("/help", "LoginWindow", "height=300,width=400,scrollbars=1,resizable=1,menubar=0,location=0,status=0,toolbar=0");
}