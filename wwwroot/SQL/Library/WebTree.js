// Node object
function Node($NodeId, $ParentId, $NodeName, $NodeUrl, $Target, $ToolTip) {
	this.$NodeId	= $NodeId;
	this.$ParentId	= $ParentId;
	this.$NodeName	= $NodeName;
	this.$NodeUrl	= $NodeUrl;
	this.$Target	= $Target;
	this.$ToolTip	= $ToolTip;
	this.icon		= null;
	this.iconOpen	= null;
	this._io = false;
	this._is = false;
	this._ls = false;
	this._hc = false;
	this._ai = 0;
	this._p;
};

// Tree object
function TreeView(treeName,folderLinks,allowSelect,autoCollapse,navigateUrl,eventName) {
	this.config = {
		$TreeName		: treeName,
		$FolderLinks	: folderLinks,
		$AllowSelect	: allowSelect,
		$UseCookies		: true,
		$ShowLines		: true,
		$ShowIcons		: true,
		$AutoCollapse	: autoCollapse,
		$InOrder		: true,
		$EventName		: eventName,
		$NavigateUrl	: navigateUrl
	}
	this.icon = {
		root			: 'Style/WebTree/root.gif',
		folder			: 'Style/WebTree/folder.gif',
		folderOpen		: 'Style/WebTree/folder.gif',
		node			: 'Style/WebTree/item.gif',
		empty			: 'Style/WebTree/empty.gif',
		line			: 'Style/WebTree/empty.gif',
		join			: 'Style/WebTree/empty.gif',
		joinBottom		: 'Style/WebTree/empty.gif',
		plus			: 'Style/WebTree/plus.gif',
		plusBottom		: 'Style/WebTree/plus.gif',
		minus			: 'Style/WebTree/minus.gif',
		minusBottom		: 'Style/WebTree/minus.gif',
		nlPlus			: 'Style/WebTree/plus.gif',
		nlMinus			: 'Style/WebTree/minus.gif'

	};
	this.obj = treeName;
	this.aNodes = [];
	this.aIndent = [];
	this.root = new Node(-1);
	this.selectedNode = null;
	this.selectedFound = false;
	this.completed = false;
};

// Adds a new node to the node array
TreeView.prototype.add = function($NodeId, $ParentId, $NodeName, $NodeUrl, $Target, $ToolTip) {
	this.aNodes[this.aNodes.length] = new Node($NodeId, $ParentId, $NodeName, $NodeUrl, $Target, $ToolTip);
};

// Open/close all nodes
TreeView.prototype.openAll = function() {
	this.oAll(true);
};
TreeView.prototype.closeAll = function() {
	this.oAll(false);
};


// Outputs the tree to the page
TreeView.prototype.toString = function() {
	var str = '<div class="TreeView">\n';
	if (document.getElementById) {
		if (this.config.$UseCookies) this.selectedNode = this.getSelected();
		str += this.addNode(this.root);
	} else str += 'Browser not supported.';
	str += '</div>';
	if (!this.selectedFound) this.selectedNode = null;
	this.completed = true;
	return str;
};

// Creates the tree structure
TreeView.prototype.addNode = function(pNode) {
	var str = '';
	var n=0;
	if (this.config.$InOrder) n = pNode._ai;
	for (n; n<this.aNodes.length; n++) {
		if (this.aNodes[n].$ParentId == pNode.$NodeId) {
			var cn = this.aNodes[n];
			cn._p = pNode;
			cn._ai = n;
			this.setCS(cn);
			if (cn._hc && !cn._io && this.config.$UseCookies) cn._io = this.isOpen(cn.$NodeId);
			if (!this.config.$FolderLinks && cn._hc) cn.$NodeUrl = null;
			if (this.config.$AllowSelect && cn.$NodeId == this.selectedNode && !this.selectedFound) {
					cn._is = true;
					this.selectedNode = n;
					this.selectedFound = true;
			}
			str += this.node(cn, n);
			if (cn._ls) break;
		}
	}
	return str;
};

// Creates the node icon, url and text
TreeView.prototype.node = function(node, nodeId) {
	var str = '<div class="TreeViewNode">' + this.indent(node, nodeId);
	if (this.config.$ShowIcons) {
		if (!node.icon) node.icon = (this.root.$NodeId == node.$ParentId) ? this.icon.root : ((node._hc) ? this.icon.folder : this.icon.node);
		if (!node.iconOpen) node.iconOpen = (node._hc) ? this.icon.folderOpen : this.icon.node;
		if (this.root.$NodeId == node.$ParentId) {
			node.icon = this.icon.root;
			node.iconOpen = this.icon.root;
		}
		str += '<img border=0 id="i' + this.obj + nodeId + '" src="' + ((node._io) ? node.iconOpen : node.icon) + '" alt="" />';
	}

	if (!this.config.$FolderLinks && node._hc && node.$ParentId != this.root.$NodeId)
	{
		str += '<a id="s' + this.obj + nodeId + '" href="javascript: ' + this.obj + '.o(' + nodeId + ');" class="node">';
	}
	else if (node.$ParentId != this.root.$NodeId)
	{
		str += '<a id="s' + this.obj + nodeId + '" class="' + ((this.config.$AllowSelect) ? ((node._is ? 'nodeSel' : 'node')) : 'node') + '" style="cursor=hand"';
		if (node.$ToolTip) 	str += ' title="' + node.$ToolTip + '"';		
		if (this.config.$AllowSelect)
		{
			str += ' onclick="javascript: ' + this.obj + '.s(' + nodeId + ');';
			if (this.config.$NavigateUrl && node.$NodeUrl) 
			{			
				str += '"';
				str += ' href="' + node.$NodeUrl + '"';
				if (node.$Target) 	str += ' target="' + node.$Target + '"';
			}
			else
			if (this.config.$EventName)
			{
				str += ' ' + this.config.$EventName + '(\'' + node.$NodeId + '\',\'' + node.$NodeName + '\',\'' + node.$NodeUrl + '\');';
				str += '"';
			}
			else
			{
				str += '"';
			}
		}
		str += '>';		
	}
	str += node.$NodeName;
	str += '</a>';
	str += '</div>';
	if (node._hc) {
		str += '<div " id="d' + this.obj + nodeId + '" class="clip" style="display:' + ((this.root.$NodeId == node.$ParentId || node._io) ? 'block' : 'none') + ';">';
		str += this.addNode(node);
		str += '</div>';
	}
	this.aIndent.pop();
	return str;
};

// Adds the empty and line icons
TreeView.prototype.indent = function(node, nodeId) {
	var str = '';
	if (this.root.$NodeId != node.$ParentId) {
		for (var n=0; n<this.aIndent.length; n++)
			str += '<img border=0 src="' + ( (this.aIndent[n] == 1 && this.config.$ShowLines) ? this.icon.line : this.icon.empty ) + '" alt="" />';
		(node._ls) ? this.aIndent.push(0) : this.aIndent.push(1);
		if (node._hc) {
			str += '<a href="javascript: ' + this.obj + '.o(' + nodeId + ');"><img border=0 id="j' + this.obj + nodeId + '" src="';
			if (!this.config.$ShowLines) str += (node._io) ? this.icon.nlMinus : this.icon.nlPlus;
			else str += ( (node._io) ? ((node._ls && this.config.$ShowLines) ? this.icon.minusBottom : this.icon.minus) : ((node._ls && this.config.$ShowLines) ? this.icon.plusBottom : this.icon.plus ) );
			str += '" alt="" /></a>';
		} else str += '<img border=0 src="' + ( (this.config.$ShowLines) ? ((node._ls) ? this.icon.joinBottom : this.icon.join ) : this.icon.empty) + '" alt="" />';
	}
	return str;
};

// Checks if a node has any children and if it is the last sibling
TreeView.prototype.setCS = function(node) {
	var lastId;
	for (var n=0; n<this.aNodes.length; n++) {
		if (this.aNodes[n].$ParentId == node.$NodeId) node._hc = true;
		if (this.aNodes[n].$ParentId == node.$ParentId) lastId = this.aNodes[n].$NodeId;
	}
	if (lastId==node.$NodeId) node._ls = true;
};

// Returns the selected node
TreeView.prototype.getSelected = function() {
	var sn = this.getCookie('cs' + this.obj);
	return (sn) ? sn : null;
};

// Highlights the selected node
TreeView.prototype.s = function($NodeId) {
	if (!this.config.$AllowSelect) return;
	var cn = this.aNodes[$NodeId];
	if (cn._hc && !this.config.$FolderLinks) return;
	if (this.selectedNode != $NodeId) {
		if (this.selectedNode || this.selectedNode==0) {
			eOld = document.getElementById("s" + this.obj + this.selectedNode);
			if (eOld!=null)
			{
				eOld.className = "node";
			}
			
		}
		eNew = document.getElementById("s" + this.obj + $NodeId);
		eNew.className = "nodeSel";
		this.selectedNode = $NodeId;
		if (this.config.$UseCookies) this.setCookie('cs' + this.obj, cn.$NodeId);
	}
};

// Toggle Open or close
TreeView.prototype.o = function($NodeId) {
	var cn = this.aNodes[$NodeId];
	this.nodeStatus(!cn._io, $NodeId, cn._ls);
	cn._io = !cn._io;
	if (this.config.$AutoCollapse) this.closeLevel(cn);
	if (this.config.$UseCookies) this.updateCookie();
};

// Open or close all nodes
TreeView.prototype.oAll = function(status) {
	for (var n=0; n<this.aNodes.length; n++) {
		if (this.aNodes[n]._hc && this.aNodes[n].$ParentId != this.root.$NodeId) {
			this.nodeStatus(status, n, this.aNodes[n]._ls)
			this.aNodes[n]._io = status;
		}
	}
	if (this.config.$UseCookies) this.updateCookie();
};

// Opens the tree to a specific node
TreeView.prototype.openTo = function(nId, bSelect, bFirst) {
	if (!bFirst) {
		for (var n=0; n<this.aNodes.length; n++) {
			if (this.aNodes[n].$NodeId == nId) {
				nId=n;
				break;
			}
		}
	}
	var cn=this.aNodes[nId];
	if (cn.$ParentId==this.root.$NodeId || !cn._p) return;
	cn._io = true;
	cn._is = bSelect;
	if (this.completed && cn._hc) this.nodeStatus(true, cn._ai, cn._ls);
	if (this.completed && bSelect) this.s(cn._ai);
	else if (bSelect) this._sn=cn._ai;
	this.openTo(cn._p._ai, false, true);
};

// Closes all nodes on the same level as certain node
TreeView.prototype.closeLevel = function(node) {
	for (var n=0; n<this.aNodes.length; n++) {
		if (this.aNodes[n].$ParentId == node.$ParentId && this.aNodes[n].$NodeId != node.$NodeId && this.aNodes[n]._hc) {
			this.nodeStatus(false, n, this.aNodes[n]._ls);
			this.aNodes[n]._io = false;
			this.closeAllChildren(this.aNodes[n]);
		}
	}
}

// Closes all children of a node
TreeView.prototype.closeAllChildren = function(node) {
	for (var n=0; n<this.aNodes.length; n++) {
		if (this.aNodes[n].$ParentId == node.$NodeId && this.aNodes[n]._hc) {
			if (this.aNodes[n]._io) this.nodeStatus(false, n, this.aNodes[n]._ls);
			this.aNodes[n]._io = false;
			this.closeAllChildren(this.aNodes[n]);		
		}
	}
}

// Change the status of a node(open or closed)
TreeView.prototype.nodeStatus = function(status, $NodeId, bottom) {
	eDiv	= document.getElementById('d' + this.obj + $NodeId);
	eJoin	= document.getElementById('j' + this.obj + $NodeId);
	if (this.config.$ShowIcons) {
		eIcon	= document.getElementById('i' + this.obj + $NodeId);
		eIcon.src = (status) ? this.aNodes[$NodeId].iconOpen : this.aNodes[$NodeId].icon;
	}
	eJoin.src = (this.config.$ShowLines)?
	((status)?((bottom)?this.icon.minusBottom:this.icon.minus):((bottom)?this.icon.plusBottom:this.icon.plus)):
	((status)?this.icon.nlMinus:this.icon.nlPlus);
	eDiv.style.display = (status) ? 'block': 'none';
};

// [Cookie] Clears a cookie
TreeView.prototype.clearCookie = function() {
	var now = new Date();
	var yesterday = new Date(now.getTime() - 1000 * 60 * 60 * 24);
	this.setCookie('co'+this.obj, 'cookieValue', yesterday);
	this.setCookie('cs'+this.obj, 'cookieValue', yesterday);
};

// [Cookie] Sets value in a cookie
TreeView.prototype.setCookie = function(cookieName, cookieValue, expires, path, domain, secure) {
	document.cookie =
		escape(cookieName) + '=' + escape(cookieValue)
		+ (expires ? '; expires=' + expires.toGMTString() : '')
		+ (path ? '; path=' + path : '')
		+ (domain ? '; domain=' + domain : '')
		+ (secure ? '; secure' : '');
};

// [Cookie] Gets a value from a cookie
TreeView.prototype.getCookie = function(cookieName) {
	var cookieValue = '';
	var posName = document.cookie.indexOf(escape(cookieName) + '=');
	if (posName != -1) {
		var posValue = posName + (escape(cookieName) + '=').length;
		var endPos = document.cookie.indexOf(';', posValue);
		if (endPos != -1) cookieValue = unescape(document.cookie.substring(posValue, endPos));
		else cookieValue = unescape(document.cookie.substring(posValue));
	}
	return (cookieValue);
};

// [Cookie] Returns ids of open nodes as a string
TreeView.prototype.updateCookie = function() {
	var str = '';
	for (var n=0; n<this.aNodes.length; n++) {
		if (this.aNodes[n]._io && this.aNodes[n].$ParentId != this.root.$NodeId) {
			if (str) str += '.';
			str += this.aNodes[n].$NodeId;
		}
	}
	this.setCookie('co' + this.obj, str);
};

// [Cookie] Checks if a node id is in a cookie
TreeView.prototype.isOpen = function($NodeId) {
	var aOpen = this.getCookie('co' + this.obj).split('.');
	for (var n=0; n<aOpen.length; n++)
		if (aOpen[n] == $NodeId) return true;
	return false;

};

// If Push and pop is not implemented by the browser
if (!Array.prototype.push) {
	Array.prototype.push = function array_push() {
		for(var i=0;i<arguments.length;i++)
			this[this.length]=arguments[i];
		return this.length;
	}
};

if (!Array.prototype.pop) {
	Array.prototype.pop = function array_pop() {
		lastElement = this[this.length-1];
		this.length = Math.max(this.length-1,0);
		return lastElement;
	}
};