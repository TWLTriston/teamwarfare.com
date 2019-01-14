	fFauxHover = function() {
		if (document.getElementById("nav")) {
			var oEl = document.getElementById("nav").getElementsByTagName("LI");
			for (var i=0; i<oEl.length; i++) {
				oEl[i].onmouseover=function() {
					this.className+=" fauxhover";
				}
				oEl[i].onmouseout=function() {
					this.className=this.className.replace(new RegExp(" fauxhover\\b"), "");
				}
			}
		}
	}
	if (window.attachEvent) window.attachEvent("onload", fFauxHover);
	
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