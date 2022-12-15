/******************** Namespaces.Grid.js  ********************/
if (!window.Active)	  	{var Active			= {}}
if (!Active.System)	  	{Active.System		= {}}
if (!Active.HTML)	  	{Active.HTML		= {}}
if (!Active.Templates)	{Active.Templates 	= {}}
if (!Active.Formats)	{Active.Formats 	= {}}
if (!Active.HTTP)		{Active.HTTP	 	= {}}
if (!Active.Text)		{Active.Text	 	= {}}
if (!Active.XML)		{Active.XML		 	= {}}
if (!Active.Controls) 	{Active.Controls	= {}}
/******************** Browsers.Gecko.js   ********************/
(function(){
	if (!window.HTMLElement) {return}
	var element = HTMLElement.prototype;
    element.__proto__ = {__proto__: element.__proto__};
    element = element.__proto__;
//	------------------------------------------------------------
	var capture = ["click",	"mousedown", "mouseup",	"mousemove", "mouseover", "mouseout" ];
    element.setCapture = function(){
        var self = this;
        var flag = false;
        this._capture = function(e){
            if (flag) {return}
            flag = true;
            var event = document.createEvent("MouseEvents");
            event.initMouseEvent(e.type,
                e.bubbles, e.cancelable, e.view, e.detail,
                e.screenX, e.screenY, e.clientX, e.clientY,
                e.ctrlKey, e.altKey, e.shiftKey, e.metaKey,
                e.button, e.relatedTarget);
            self.dispatchEvent(event);
            flag = false;
        };
        for (var i=0; i<capture.length; i++) {
            window.addEventListener(capture[i], this._capture, true);
        }
    };
    element.releaseCapture = function(){
        for (var i=0; i<capture.length; i++) {
            window.removeEventListener(capture[i], this._capture, true);
        }
        this._capture = null;
    };
//	------------------------------------------------------------
	element.attachEvent = function (name, handler) {
		if (typeof handler != "function") {return}
		var nsName = name.replace(/^on/, "");
		var nsHandler = function(event){
			window.event = event;
			handler();
			window.event = null;
		};
		handler[name] = nsHandler;
		this.addEventListener(nsName, nsHandler, false);
	};
	element.detachEvent = function (name, handler) {
		if (typeof handler != "function") {return}
		var nsName = name.replace(/^on/, "");
		this.removeEventListener(nsName, handler[name], false);
		handler[name] = null;
	};
//	------------------------------------------------------------
	var getClientWidth = function(){
		return this.offsetWidth - 20;
	};
	var getClientHeight = function(){
		return this.offsetHeight - 20;
	};
	element.__defineGetter__("clientWidth", getClientWidth);
	element.__defineGetter__("clientHeight", getClientHeight);
//	------------------------------------------------------------
	var getRuntimeStyle = function(){
		return this.style;
	};
	element.__defineGetter__("runtimeStyle", getRuntimeStyle);
//	------------------------------------------------------------
	var cs = ComputedCSSStyleDeclaration.prototype;
    cs.__proto__ = {__proto__: cs.__proto__};
    cs = cs.__proto__;
	cs.__defineGetter__("paddingTop", function(){return this.getPropertyValue("padding-top")});
	var getCurrentStyle = function(){
		return document.defaultView.getComputedStyle(this, "");
	};
	element.__defineGetter__("currentStyle", getCurrentStyle);
//	------------------------------------------------------------
	var setOuterHtml = function(s){
	   var range = this.ownerDocument.createRange();
	   range.setStartBefore(this);
	   var fragment = range.createContextualFragment(s);
	   this.parentNode.replaceChild(fragment, this);
	};
	element.__defineSetter__("outerHTML", setOuterHtml);
})();
//	------------------------------------------------------------
(function(){
	if (!window.Event) {return}
	var event = Event.prototype;
    event.__proto__ = {__proto__: event.__proto__};
	event = event.__proto__;
	if (!event) {return}
//	------------------------------------------------------------
	var getSrcElement = function(){
		return (this.target.nodeType==3) ? this.target.parentNode : this.target;
	};
	event.__defineGetter__("srcElement", getSrcElement);
//	------------------------------------------------------------
	var setReturnValue = function(value){
		if (!value) {this.preventDefault()}
	};
	event.__defineSetter__("returnValue", setReturnValue);
})();
//	------------------------------------------------------------
(function(){
	if (!window.CSSStyleSheet){return}
	var stylesheet = CSSStyleSheet.prototype;
    stylesheet.__proto__ = {__proto__: stylesheet.__proto__};
    stylesheet = stylesheet.__proto__;
	stylesheet.addRule = function(selector, rule){
		this.insertRule(selector + "{" + rule + "}", this.cssRules.length);
	};
	stylesheet.__defineGetter__("rules", function(){return this.cssRules});
})();
//	------------------------------------------------------------
(function(){
	if (!window.XMLHttpRequest) {return}
	var ActiveXObject = function(type) {
		ActiveXObject[type](this);
	};
	ActiveXObject["MSXML2.DOMDocument"] = function(obj){
		obj.setProperty = function(){};
		obj.load = function(url){
			var xml = this;
			var async = this.async ? true : false;
			var request = new XMLHttpRequest();
			request.open("GET", url, async);
			request.overrideMimeType("text/xml");
			if (async) {
				request.onreadystatechange = function(){
					xml.readyState = request.readyState;
					if (request.readyState == 4 ) {
						xml.documentElement = request.responseXML.documentElement;
						xml.firstChild = xml.documentElement; 
						request.onreadystatechange = null;
					}
					if (xml.onreadystatechange) {xml.onreadystatechange()}
				}
			}
			this.parseError = {errorCode: 0, reason: "Emulation"};
			request.send(null);
			this.readyState = request.readyState;
			if (request.responseXML && !async) {
				this.documentElement = request.responseXML.documentElement;
				this.firstChild = this.documentElement; 
			}
		}
	};
	ActiveXObject["MSXML2.XMLHTTP"] = function(obj){
		obj.open = function(method, url, async){
			this.request = new XMLHttpRequest();
			this.request.open(method, url, async);
		};
		obj.send = function(data){
			this.request.send(data);
		};
		obj.setRequestHeader = function(name, value){
			this.request.setRequestHeader(name, value);
		};
		obj.__defineGetter__("readyState", function(){return this.request.readyState});
		obj.__defineGetter__("responseXML", function(){return this.request.responseXML});
		obj.__defineGetter__("responseText", function(){return this.request.responseText});
	};
//	window.ActiveXObject = ActiveXObject;
})();
//	------------------------------------------------------------
(function(){
	if (!window.XPathEvaluator) {return}
	var xpath = new XPathEvaluator();
	var element = Element.prototype;
    element.__proto__ = {__proto__: element.__proto__};
    element = element.__proto__;
	var attribute = Attr.prototype;
    attribute.__proto__ = {__proto__: attribute.__proto__};
    attribute = attribute.__proto__;
	var txt = Text.prototype;
    txt.__proto__ = {__proto__: txt.__proto__};
    txt = txt.__proto__;
	var doc = Document.prototype;
    doc.__proto__ = {__proto__: doc.__proto__};
    doc = doc.__proto__;
	doc.loadXML = function(text){
		var parser = new DOMParser;
		var newDoc = parser.parseFromString(text, "text/xml");
		this.replaceChild(newDoc.documentElement, this.documentElement);
	};
	doc.setProperty = function(name, value){
		if(name=="SelectionNamespaces"){
			namespaces = {};
			var a = value.split(" xmlns:");
			for (var i=1;i<a.length;i++){
				var s = a[i].split("=");
				namespaces[s[0]] = s[1].replace(/\"/g, "");
			}
			this._ns = {
				lookupNamespaceURI : function(prefix){return namespaces[prefix]}
			}
		}
	};
	doc._ns = {
		lookupNamespaceURI : function(){return null}
	};
	doc.selectNodes = function (path) {
	   var result = xpath.evaluate(path, this, this._ns, 7, null);
	   var i, nodes = [];
	   for (i=0; i<result.snapshotLength; i++) {nodes[i]=result.snapshotItem(i)}
	   return nodes;
	};
	doc.selectSingleNode = function (path) {
	   return xpath.evaluate(path, this, this._ns, 9, null).singleNodeValue;
	};
	element.selectNodes = function (path) {
	   var result = xpath.evaluate(path, this, this.ownerDocument._ns, 7, null);
	   var i, nodes = [];
	   for (i=0; i<result.snapshotLength; i++) {nodes[i]=result.snapshotItem(i)}
	   return nodes;
	};
	element.selectSingleNode = function (path) {
	   return xpath.evaluate(path, this, this.ownerDocument._ns, 9, null).singleNodeValue;
	};
	element.__defineGetter__("text", function(){
		var i, a=[], nodes = this.childNodes, length = nodes.length;
		for (i=0; i<length; i++){a[i] = nodes[i].text}
		return a.join("");
	});
	attribute.__defineGetter__("text", function(){return this.nodeValue});
	txt.__defineGetter__("text", function(){return this.nodeValue});
})();
/******************** System.Object.js    ********************/
Active.System.Object = function(){};
/*
	var Active is an object, the root of the hierarchy. Active.System is
	also an object (System is a property of Active) Active.System.Object
	is a function (Object is a method of System). To be precise, it is
	not just a function, it is a constructor function, which is used to
	create objects of type Active.System.Object (like this: var obj = new
	Active.System.Object;).
*/
Active.System.Object.subclass = function(){
/*
	We are creating a method 'subclass' of the Active.System.Object.
	Again, Active.System.Object is a constructor function, which
	represents a class, not an object instance. It is OK to create a
	method or a property of a function because functions themselves
	behave like objects.
*/
	var constructor = function(){this.init()};
/*
	Our 'subclass' method should return a constructor function, which
	we create here. Because the constructor is created automatically,
	the actual object initialization code should be somewhere else,
	i.e. in a special 'init' method. So each object constructor just
	calls 'init' method of a newly created object. Here 'this' keyword
	refers to the newly created object, when subclass constructor runs.
*/
	for (var i in this) {constructor[i] = this[i]}
/*
	This code copies all properties and methods from the base class
	constructor to the derived class constructor. Note, it is NOT object
	properties, it is constructor function properties. Keyword 'this'
	refers to the base class constructor, i.e. Active.System.Object
	function.
*/
	constructor.prototype = new this();
/*
	The 'prototype' property of the constructor of the derived class
	should point to the base class object instance. Here we create a new
	instance of the base class by calling the base class constructor
	function with the keyword 'new'. Again, 'this' refers to the base
	class constructor, i.e. Active.System.Object function.
*/
	constructor.superclass = this;
/*
	We also create special 'superclass' property, which provides quick
	access to the base class constructor from within the derived class.
	It is very useful when you want to overload a method in the derived
	class but still be able to call the base class implementation of the
	same method.
*/
	return constructor;
};
Active.System.Object.handle = function(error){
	throw(error);
};
Active.System.Object.create = function(){
/****************************************************************
	Generic base class - root of the ActiveWidgets class hierarchy.
*****************************************************************/
	var obj = this.prototype;
/****************************************************************
	Creates an object clone.
	@return		A new object.
	The clone function creates a fast copy of the object. Instead of
	physically copying each property and method of the source object -
	it creates a clone as a ‘subclass’ of the source object, i.e.
	properties and methods  are inherited from the source object into
	the clone.
	Note that the clone continues to be dependent on the source
	object. Changes in the source object property or method will
	affect all the clones unless this property is already overwritten
	in the clone object itself.
*****************************************************************/
	obj.clone = function(){
		if (this._clone.prototype!==this) {
			this._clone = function(){this.init()};
			this._clone.prototype = this;
		}
		return new this._clone();
	};
	obj._clone = function(){};
/****************************************************************
	Initializes the object.
	@remarks
	This method normaly contains all object initialization code
	(instead of the constructor function).	Constructor function is
	the same for all objects and only contains object.init() call.
*****************************************************************/
	obj.init = function(){
		// overload
	};
/****************************************************************
	Handles exceptions in the ActiveWidgets methods.
	@param	error (Error) Error object.
	The default error handler just throws the same exception to the
	next level. Overload this function to add your own diagnostics
	and error logging.
*****************************************************************/
 	obj.handle = function(error){
		throw(error);
	};
/****************************************************************
	Calls a method after a specified time interval has elapsed.
	@param	handler (Function) Method to call.
	@param	delay (Number) Time interval in milliseconds.
	@return An identifier that can be used with window.clearTimeout
			to cancel the current method call.
	This method has the same effect as window.setTimeout except that
	the function will be evaluated not as a global function but
	as a method of the current object.
*****************************************************************/
	obj.timeout = function(handler, delay){
		var self = this;
		var wrapper = function(){handler.call(self)};
		return window.setTimeout(wrapper, delay ? delay : 0);
	};
/****************************************************************
	Converts object to string.
	@return Text or HTML representation of the object.
	This method is overloaded in ActiveWidgets subclasses.
*****************************************************************/
 	obj.toString = function(){
		return "";
	};
};
Active.System.Object.create();
/******************** System.Model.js     ********************/
Active.System.Model = Active.System.Object.subclass();
Active.System.Model.create = function(){
/****************************************************************
	Generic data model class.
*****************************************************************/
	var obj = this.prototype;
	var join = function(){
		var i, s = arguments[0];
		for (i=1; i<arguments.length; i++){s += arguments[i].substr(0,1).toUpperCase() + arguments[i].substr(1)}
		return s;
	};
/****************************************************************
	Creates a new property.
	@param	name	(String) Property name.
	@param	value	(String) Default property value.
*****************************************************************/
	obj.defineProperty = function(name, value){
		var _getProperty = join("get", name);
		var _setProperty = join("set", name);
		var _property = "_" + name;
		var getProperty = function(){
			return this[_property];
		};
		this[_setProperty] = function(value){
			if(typeof value == "function"){
				this[_getProperty] = value;
			}
			else {
				this[_getProperty] = getProperty;
				this[_property] = value;
			}
		};
		this[_setProperty](value);
	};
	var get = {};
	var set = {};
/****************************************************************
	Returns property value.
	@param	name	(String) Property name.
	@return Property value.
*****************************************************************/
	obj.getProperty = function(name, a, b, c){
		if (!get[name]) {get[name] = join("get", name)}
		return this[get[name]](a, b, c);
	};
/****************************************************************
	Sets property value.
	@param	name	(String) Property name.
	@param	value	(String) Property value.
*****************************************************************/
	obj.setProperty = function(name, value, a, b, c){
		if (!set[name]) {set[name] = join("set", name)}
		return this[set[name]](value, a, b, c);
	};
/****************************************************************
	Indicates whether the data is available.
*****************************************************************/
	obj.isReady = function(){
		return true;
	};
};
Active.System.Model.create();
/******************** System.Format.js    ********************/
Active.System.Format = Active.System.Object.subclass();
Active.System.Format.create = function(){
/****************************************************************
	Generic data formatting class.
*****************************************************************/
	var obj = this.prototype;
/****************************************************************
	Transforms the primitive value into the readable text.
	@param	value	(Any) Primitive value.
	@return		Readable text.
*****************************************************************/
	obj.valueToText = function(value){
		return value;
	};
/****************************************************************
	Transforms the wire data into the primitive value.
	@param	data	(String) Wire data.
	@return		Primitive value.
*****************************************************************/
	obj.dataToValue = function(data){
		return data;
	};
/****************************************************************
	Transforms the wire data into the readable text.
	@param	data	(String) Wire data.
	@return		Readable text.
*****************************************************************/
	obj.dataToText = function(data){
		var value = this.dataToValue(data);
		return this.valueToText(value);
	};
/****************************************************************
	Specifies the text to be returned in case of error.
	@param	text	(String) Error text.
*****************************************************************/
	obj.setErrorText = function(text){
		this._textError = text;
	};
/****************************************************************
	Specifies the value to be returned in case of error.
	@param	value	(Any) Error value.
*****************************************************************/
	obj.setErrorValue = function(value){
		this._valueError = value;
	};
	obj.setErrorText("#ERR");
	obj.setErrorValue(NaN);
};
Active.System.Format.create();
/******************** System.Html.js      ********************/
Active.System.HTML = Active.System.Object.subclass();
Active.System.HTML.create = function(){
/****************************************************************
	Generic base class for building and manipulating HTML markup.
	Objects, which  have visual representation, are most likely
	subclasses of this generic HTML class. It provides a set of
	functions to define attributes, inline styles, stylesheet
	selectors, DOM events and inner HTML content either as static
	properties or calls to the object’s methods. Direct or implicit
	call to ‘toString’ method returns properly formatted HTML
	markup string, which can be used in document.write() call or
	assigned to the page innerHTML property.
	The two-way linking between original javascript object and
	it’s DOM counterpart is maintained through the use of unique ID for
	each object. This allows forwarding DOM events back to the
	proper javascript master object and, if necessary, updating
	the correct piece of HTML on the page.
*****************************************************************/
	var obj = this.prototype;
/****************************************************************
	Sets HTML tag for the object.
	@param	tag (String) The new tag.
	By default each HTML object is a DIV tag. This function allows
	to change the tag string.
	@example
	obj.setTag("SPAN");
*****************************************************************/
	obj.setTag = function(tag){
		this._tag = tag;
	};
/****************************************************************
	Returns HTML tag for the object.
	@return	HTML tag string
*****************************************************************/
	obj.getTag = function(){
		return this._tag;
	};
	obj._tag = "div";
/****************************************************************
	Initializes the object.
*****************************************************************/
	obj.init = function(){
		if (this.$owner) {return}
		if (this._parent) {return}
		this._id = "tag" + this.all.id++;
		this.all[this._id] = this;
	};
/****************************************************************
	Returns unique ID for the object.
	@return	Unique ID string.
*****************************************************************/
	obj.getId = function(){
		return this._id;
	};
	obj._id = "";
	obj.all = Active.System.all = {id:0};
/****************************************************************
	Sets ID string for an element.
	@param	id (String) New ID.
*****************************************************************/
	obj.setId = function(id){
		this._id = id;
		this.all[this._id] = this;
	};
/****************************************************************
	Returns a reference to the HTML element.
	@return Reference to the HTML element
	This function returns null if it is called before writing the
	object to the page.
*****************************************************************/
	obj.element = function(){
		var i, docs = this._docs, id = this.getId(), e;
		for(i=0; i<docs.length; i++) {
			e = docs[i].getElementById(id);
			if(e) {return e}
		}
	};
	obj._docs = [document];
/****************************************************************
	Returns CSS selector.
	@param	name (String) Selector name.
	@return	Selector value.
*****************************************************************/
	obj.getClass = function(name){
		var param = "_" + name + "Class";
		var value = this[param];
		return typeof(value)=="function" ? value.call(this) : value;
	};
/****************************************************************
	Sets CSS selector.
	@param	name (String) Selector name.
	@param	value (String/Function) Selector value.
	The selector string is composed from the three parts - the prefix
	('active'),	the name and the value, separated by the '-' character.
	Normally the object class string consists of several selectors
	separated by space.
	Selector values are stored and inherited separately within the
	object. This function allows easy access to single selector
	value without parsing the whole class string.
	The following example adds 'active-template-list' stylesheet
	selector to the object class.
	@example
	obj.setClass("template", "list");
*****************************************************************/
	obj.setClass = function(name, value){
		var element = this.element();
		if (element) {
			var v = (typeof(value)=="function") ? value.call(this) : value;
			element.className = element.className.replace(new RegExp("(active-" + name + "-\\w+ |$)"), " active-" + name + "-" + v + " ");
			if (this.$index !== "") {return} 
		}
		if (this.data) {return} 

		var param = "_" + name + "Class";
		if (this[param]==null) {this._classes += " " + name}
		this[param] = value;
		this._outerHTML = "";
	};
/****************************************************************
	Updates CSS selectors string for an element.
*****************************************************************/
	obj.refreshClasses = function(){
		var element = this.element();
		if (!element) {return}
		var s = "", classes = this._classes.split(" ");
		for (var i=1; i<classes.length; i++){
			var name = classes[i];
			var value = this["_" + name + "Class"];
			if (typeof(value)=="function") {
				value = value.call(this);
			}
			s += "active-" + name + "-" + value + " ";
		}
		element.className = s + this.$browser;
	};
	obj._classes = "";
/****************************************************************
	Returns inline CSS attribute.
	@param	name (String) CSS attribute name.
	@return	CSS attribute value.
*****************************************************************/
	obj.getStyle = function(name){
		var param = "_" + name + "Style";
		var value = this[param];
		return typeof(value)=="function" ? value.call(this) : value;
	};
/****************************************************************
	Sets inline CSS attribute.
	@param	name (String) CSS attribute name.
	@param	value (String/Function) CSS attribute value.
*****************************************************************/
	obj.setStyle = function(name, value){
		var element = this.element();
		if (element) {element.style[name] = value}
		if (this.data) {return} 

		var param = "_" + name + "Style";
		if (this[param]==null) {this._styles += " " + name}
		this[param] = value;
		this._outerHTML = "";
	};
	obj._styles = "";
/****************************************************************
	Returns HTML attribute.
	@param	name (String) HTML attribute name.
	@return	HTML attribute value.
*****************************************************************/
	obj.getAttribute = function(name){
		try {
			var param = "_" + name + "Attribute";
			var value = this[param];
			return typeof(value)=="function" ? value.call(this) : value;
		}
		catch(error){
			this.handle(error);
		}
	};
/****************************************************************
	Sets HTML attribute.
	@param	name (String) HTML attribute name.
	@param	value (String/Function) HTML attribute value.
*****************************************************************/
	obj.setAttribute = function(name, value){
		try {
			var param = "_" + name + "Attribute";
			if (typeof this[param] == "undefined") {this._attributes += " " + name}
			if (specialAttributes[name] && (typeof value == "function")){
				this[param] = function(){return value.call(this) ? true : null};
			}
			else {
				this[param] = value;
			}
			this._outerHTML = "";
		}
		catch(error){
			this.handle(error);
		}
	};
	obj._attributes = "";
	var specialAttributes = {
		checked	  : true,
		disabled  : true,
		hidefocus : true,
		readonly  : true };
/****************************************************************
	Returns HTML event handler.
	@param	name (String) HTML event name.
	@return	HTML event handler.
*****************************************************************/
	obj.getEvent = function(name){
		try {
			var param = "_" + name + "Event";
			var value = this[param];
			return value;
		}
		catch(error){
			this.handle(error);
		}
	};
/****************************************************************
	Sets HTML event handler.
	@param	name (String) HTML event name.
	@param	value (String/Function) HTML event handler.
*****************************************************************/
	obj.setEvent = function(name, value){
		try {
			var param = "_" + name + "Event";
			if (this[param]==null) {this._events += " " + name}
			this[param] = value;
			this._outerHTML = "";
		}
		catch(error){
			this.handle(error);
		}
	};
	obj._events = "";
/****************************************************************
	Returns static HTML content.
	@param	name (String) content name.
	@return	content object or function.
*****************************************************************/
	obj.getContent = function(name){
		try {
			var split = name.match(/^(\w+)\W(.+)$/);
			if (split) {
				var ref = this.getContent(split[1]);
				return ref.getContent(split[2]);
			}
			else {
				var param = "_" + name + "Content";
				var value = this[param];
				if ((typeof value == "object") && (value._parent != this)) {
					value = value.clone();
					value._parent = this; 
					value._id = this._id + "/" + name; 
					this[param] = value;
				}
				return value;
			}
		}
		catch(error){
			this.handle(error);
		}
	};
/****************************************************************
	Sets static HTML content.
	@param	name (String) content name.
	@param	value (Object/String/Function) static content.
*****************************************************************/
	obj.setContent = function(name, value){
		try {
			if (arguments.length==1) { // assigning array or single function
				
				this._content = "";
				if (typeof name == "object") {
					for (var i in name)	{
						if (typeof(i) == "string") {
							this.setContent(i, name[i]);
						}
					}
				}
				else {
					this.setContent("html", name);
				}
			}
			else {
				var split = name.match(/^(\w+)\W(.+)$/);
				if (split) {
					var ref = this.getContent(split[1]);
					ref.setContent(split[2], value);
					this._innerHTML = "";
					this._outerHTML = "";
				}
				else {
					var param = "_" + name + "Content";
					if (this[param]==null) {this._content += " " + name}
					if (typeof value == "object") {
						value._parent = this; 
						value._id = this._id + "/" + name; 
					}
					this[param] = value;
					this._innerHTML = "";
					this._outerHTML = "";
				}
			}
		}
		catch(error){
			this.handle(error);
		}
	};
	obj._content = "";
	obj.$index = ""; 

//	------------------------------------------------------------
	var getParamStr = function(i){return "{#" + i + "}"};
//	------------------------------------------------------------
	obj.innerHTML = function(){
		//	Returns 'inner HTML' string for an object.
		try {
			// just return cached value if available
			if (this._innerHTML) {return this._innerHTML}
			this._innerParamLength = 0;
			var i, j, name, value, param1, param2, html, item, s = "";
			var content = this._content.split(" ");
			for (i=1; i<content.length; i++){
				name = content[i];
				value = this["_" + name + "Content"];
				if (typeof(value)=="function") {
					param = getParamStr(this._innerParamLength++);
					this[param] = value;
					s += param;
				}
				else if (typeof(value)=="object"){
					item = value;
					html = item.outerHTML().replace(/\{id\}/g, "{id}/" + name);
					for (j=item._outerParamLength-1; j>=0; j--){
						param1 = getParamStr(j);
						param2 = getParamStr(this._innerParamLength + j);
						if (param1 != param2) {html = html.replace(param1, param2)}
						this[param2] = item[param1];
					}
					this._innerParamLength += item._outerParamLength;
					s += html;
				}
				else {
					s += value;
				}
			}
			this._innerHTML = s;
			return s;
		}
		catch(error){
			this.handle(error);
		}
	};
//	------------------------------------------------------------
	obj.outerHTML = function(){
		//	Returns 'outer HTML' string for an object.
		try {
			// just return cached value if available
			if (this._outerHTML) {return this._outerHTML}
			// build inner HTML first
			var innerHTML = this.innerHTML();
			// reset param count
			this._outerParamLength = this._innerParamLength;
			// elementless templates
			if (!this._tag) {return innerHTML}
			var i, tmp, name, value, param;
			var html = "<" + this._tag + " id=\"{id}\"";
			tmp = "";
			var classes = this._classes.split(" ");
			for (i=1; i<classes.length; i++){
				name = classes[i];
				value = this["_" + name + "Class"];
				if (typeof(value)=="function") {
					param = getParamStr(this._outerParamLength++);
					this[param] = value;
					value = param;
				}
				tmp += "active-" + name + "-" + value + " ";
			}
			if (tmp) {html += " class=\"" + tmp + this.$browser + "\""}
			tmp = "";
			var styles = this._styles.split(" ");
			for (i=1; i<styles.length; i++){
				name = styles[i];
				value = this["_" + name + "Style"];
				if (typeof(value)=="function") {
					param = getParamStr(this._outerParamLength++);
					this[param] = value;
					value = param;
				}
				tmp += name + ":" + value + ";";
			}
			if (tmp) {html += " style=\"" + tmp + "\""}
			tmp = "";
			var attributes = this._attributes.split(" ");
			for (i=1; i<attributes.length; i++){
				name = attributes[i];
				value = this["_" + name + "Attribute"];
				if (typeof(value)=="function") {
					param = getParamStr(this._outerParamLength++);
					this[param] = value;
					value = param;
				}
				else if (specialAttributes[name] && !value ){
					value = null;
				}
				if (value !== null ){
					tmp += " " + name + "=\"" + value + "\"";
				}
			}
			html += tmp;
			tmp = "";
			var events = this._events.split(" ");
			for (i=1; i<events.length; i++){
				name = events[i];
				value = this["_" + name + "Event"];
				if (typeof(value)=="function") {
					value = "dispatch(event, this)";
				}
				tmp += " " + name + "=\"" + value + "\"";
			}
			html += tmp;
			html += ">" + innerHTML + "</" + this._tag + ">";
			// save the result in cache and return
			this._outerHTML = html;
			return html;
		}
		catch(error){
			this.handle(error);
		}
	};
/****************************************************************
	Returns HTML markup string for the object.
	@return	HTML string.
	Direct or implicit
	call to ‘toString’ method returns properly formatted HTML
	markup string, which can be used in document.write() call or
	assigned to the page innerHTML property.
*****************************************************************/
	obj.toString = function(){
		try {
			var i, s = this._outerHTML;
			if (!s) {s = this.outerHTML()}
			s = s.replace(/\{id\}/g, this.getId());
			var max = this._outerParamLength;
			for (i=0; i<max; i++){
				var param = "{#" + i + "}";
				var value = this[param]();
				if (value === null ){
					value = "";
					param = specialParams[i];
					if (!param) {param = getSpecialParamStr(i);}
				}
				s = s.replace(param, value);
			}
			return s;
		}
		catch(error){
			this.handle(error);
		}
	};
	var specialParams = [];
	function getSpecialParamStr(i){return (specialParams[i] = new RegExp("[\\w\\x2D]*=?:?\\x22?\\{#" + i + "\\}[;\\x22]?"));}
/****************************************************************
	Updates HTML on the page.
*****************************************************************/
	obj.refresh = function(){
		try {
			var element = this.element();
			if (element) {element.outerHTML = this.toString()}
		}
		catch(error){
			this.handle(error);
		}
	};
//	------------------------------------------------------------
	obj.$browser = "";
	if (window.__defineGetter__) {obj.$browser = "gecko"}
	if (navigator.userAgent.match("Opera")){obj.$browser = "opera"}
	if (navigator.userAgent.match("Konqueror")){obj.$browser = "khtml"}
	if (navigator.userAgent.match("KHTML")){obj.$browser = "khtml"}
};
Active.System.HTML.create();
// -------------------------------------------------------------------------
var dispatch = function(event, element){
	var parts = element.id.split("/");
	var tag = parts[0].split(".");
	var obj = Active.System.all[tag[0]];
	var type = "_on" + event.type + "Event";
	var i;
	for (i=1; i<tag.length; i++){
		var params = tag[i].split(":");
		obj = obj.getTemplate.apply(obj, params);
	}
	var target = obj;
	for (i=1; i<parts.length; i++){
		target = target.getContent(parts[i]);
	}
	if (window.HTMLElement) {window.event = event}
	target[type].call(obj, event); 

	if (window.HTMLElement) {window.event = null}
	return;
};
// -------------------------------------------------------------------------
var mouseover = function(element, name){
	try {
		element.className += " " + name;
	}
	catch(error){
		//	ignore errors
	}
};
var mouseout = function(element, name){
	try {
		element.className = element.className.replace(RegExp(" " + name, "g"), "");
	}
	catch(error){
		//	ignore errors
	}
};
// -------------------------------------------------------------------------
/******************** System.Template.js  ********************/
Active.System.Template = Active.System.HTML.subclass();
Active.System.Template.create = function(){
/****************************************************************
	Generic HTML template class. Template is a re-usable HTML
	fragment aimed to produce markup as part of a larger
	object (control).
	Template can either be a simple element or a complex HTML structure
	and may include calls to other templates as part of the output.
	Templates can access properties of the parent control,
	so the template output will be different depending on
	the control's data. Templates can also accept parameters
	allowing to generate lists or tables of data with the
	single instance of the template.
*****************************************************************/
	var obj = this.prototype;
	var _super = this.superclass.prototype;
	var _pattern = /^(\w+)\W(.+)$/;
	var join = function(){
		var i, s = arguments[0];
		for (i=1; i<arguments.length; i++){s += arguments[i].substr(0,1).toUpperCase() + arguments[i].substr(1)}
		return s;
	};
/****************************************************************
	Retrieves the value of the property.
	@param	name	(String) Property name.
	@return			Property value.
*****************************************************************/
	obj.getProperty = function(name, a, b, c){
		if (name.match(_pattern)) {
			var getProperty = join("get", RegExp.$1, "property");
			if (this[getProperty]) {return this[getProperty](RegExp.$2, a, b, c)}
		}
	};
/****************************************************************
	Assignes the new value to the property.
	@param	name	(String) Property name.
	@param	value	(Any) Property value.
*****************************************************************/
	obj.setProperty = function(name, value, a, b, c){
		if (name.match(_pattern)) {
			var setProperty = join("set", RegExp.$1, "property");
			if (this[setProperty]) {return this[setProperty](RegExp.$2, value, a, b, c)}
		}
	};
/****************************************************************
	Returns the data model object. For a built-in model this method
	will create a temporary proxy attached to the template.
	@param	name	(String) Name of the data model.
	@return			A data model object.
*****************************************************************/
	obj.getModel = function(name){
		var getModel = join("get", name, "model");
		return this[getModel]();
	};
/****************************************************************
	Sets the external data model.
	@param	name	(String) Name of the data model.
	@param	model	(Object) Data model object.
*****************************************************************/
	obj.setModel = function(name, model){
		var setModel = join("set", name, "model");
		return this[setModel](model);
	};
/****************************************************************
	Creates a link to the new content template.
	@param	name	(String) Template name.
	@param	template	(Object) Template object.
*****************************************************************/
	obj.defineTemplate = function(name, template){
		var ref = "_" + name + "Template";
		var get = join("get", name, "template");
		var set = join("set", name, "template");
		var getDefault = join("default", name, "template");
		var name1 = "." + name;
		var name2 = "." + name + ":";
		this[get] = this[getDefault] = function(index){
			if (typeof(this[ref])=="function") {
				return this[ref].call(this, index);
			}
			if (this[ref].$owner != this) {this[set](this[ref].clone())}
			if (typeof(index)=="undefined") {
				this[ref]._id = this._id + name1;
			}
			else {
				this[ref]._id = this._id + name2 + index;
			}
			this[ref].$index = index;
			return this[ref];
		};
		obj[get] = function(a, b, c){
			return this.$owner[get](a, b, c);
		};
		obj[set] = function(template){
			this[ref] = template;
			if (template) {
				template.$owner = this; 
			}
		};
		this[set](template);
	};
/****************************************************************
	Returns the template object.
	@param	name	(String) Template name.
	@return			Template object.
*****************************************************************/
	obj.getTemplate = function(name){
		if (name.match(_pattern)) {
			var get = join("get", RegExp.$1, "template");
			arguments[0] = RegExp.$2;
			var template = this[get]();
			return template.getTemplate.apply(template, arguments);
		}
		else {
			get = join("get", name, "template");
			var i, args = [];
			for(i=1; i<arguments.length; i++) {args[i-1]=arguments[i]}
			return this[get].apply(this, args);
		}
	};
/****************************************************************
	Sets the template.
	@param	name	(String) Template name.
	@param	template (Object) Template object.
*****************************************************************/
	obj.setTemplate = function(name, template, index){
		if (name.match(_pattern)) {
			var get = join("get", RegExp.$1, "template");
			var n = RegExp.$2;
			this[get]().setTemplate(n, template, index);
		}
		else {
			var set = join("set", name, "template");
			this[set](template, index);
		}
	};
/****************************************************************
	Returns the action handler.
	@param	name	(String) Action name.
	@return 		Action handler.
*****************************************************************/
	obj.getAction = function(name){
		return this["_" + name + "Action"];
	};
/****************************************************************
	Sets the action handler.
	@param	name	(String) Action name.
	@param	value	(Function) Action handler.
*****************************************************************/
	obj.setAction = function(name, value){
		this["_" + name + "Action"] = value;
	};
/****************************************************************
	Runs the action.
	@param	name	(String) Action name.
	@param	source	(Object) Action source.
*****************************************************************/
	obj.action = function(name, source, a, b, c){
		if (typeof source == "undefined") {source = this}
		var action = this["_" + name + "Action"];
		if (typeof(action)=="function") {action.call(this, source, a, b, c)}
		else if (this.$owner) {this.$owner.action(name, source, a, b, c)}
	};
};
Active.System.Template.create();
/******************** System.Control.js   ********************/
Active.System.Control = Active.System.Template.subclass();
Active.System.Control.create = function(){
/****************************************************************
	Generic user interface control class. Control is a screen element,
	which can have focus and responds to the keyboard or mouse commands.
	Typical control has a set of built-in or external data models
	and may also contain additional presentation templates.
*****************************************************************/
	var obj = this.prototype;
	var _super = this.superclass.prototype;
	var _pattern = /^(\w+)\W(.+)$/;
	var join = function(){
		var i, s = arguments[0];
		for (i=1; i<arguments.length; i++){s += arguments[i].substr(0,1).toUpperCase() + arguments[i].substr(1)}
		return s;
	};
	obj.setEvent("oncontextmenu", "return false");
	obj.setEvent("onselectstart", "return false");
/****************************************************************
	Creates a new data model.
	@param	name	(String) New data model name.
*****************************************************************/
	obj.defineModel = function(name){
		var external = "_" + name + "Model";
		var defineProperty = join("define", name, "property");
		var definePropertyArray = join("define", name, "property", "array");
		var getProperty = join("get", name, "property");
		var setProperty = join("set", name, "property");
		var get = {};
		var set = {};
		var getModel = join("get", name, "model");
		var setModel = join("set", name, "model");
		var updateModel = join("update", name, "model");
//		------------------------------------------------------------
		this[defineProperty] = function(property, defaultValue){
			var _getProperty = join("get", name, property);
			var _setProperty = join("set", name, property);
			var _property = "_" + join(name, property);
			var getPropertyMethod = function(){
				return this[_property];
			};
			this[_getProperty] = getPropertyMethod;
			this[_setProperty] = function(value){
				if(typeof value == "function"){
					this[_getProperty] = value;
				}
				else {
					if (this[_getProperty] !== getPropertyMethod) {this[_getProperty] = getPropertyMethod}
					this[_property] = value;
				}
				this[updateModel](property);
			};
			this[_setProperty](defaultValue);
		};
//		------------------------------------------------------------
		this[getProperty] = function(property, a, b, c){
			try {
				if (this[external]) {return this[external].getProperty(property, a, b, c)}
				if (!get[property]) {get[property] = join("get", name, property)}
				return this[get[property]](a, b, c);
			}
			catch(error){
				return this.handle(error);
			}
		};
//		------------------------------------------------------------
		this[setProperty] = function(property, value, a, b, c){
			try {
				if (this[external]) {return this[external].setProperty(property, value, a, b, c)}
				if (!set[property]) {set[property] = join("set", name, property)}
				return this[set[property]](value, a, b, c);
			}
			catch(error){
				return this.handle(error);
			}
		};
//		------------------------------------------------------------
		_super[getProperty] = function(property, a, b, c){
			if (this[external]) {return this[external].getProperty(property, a, b, c)}
			return this.$owner[getProperty](property, a, b, c);
		};
		_super[setProperty] = function(property, value, a, b, c){
			if (this[external]) {return this[external].setProperty(property, value, a, b, c)}
			return this.$owner[setProperty](property, value, a, b, c);
		};
//		------------------------------------------------------------
		this[definePropertyArray] = function(property, defaultValue){
			var _getProperty = join("get", name, property);
			var _setProperty = join("set", name, property);
			var _getArray = join("get", name, property + "s");
			var _setArray = join("set", name, property + "s");
			var _array = "_" + join(name, property + "s");
			var _getCount = join("get", name, "count");
			var _setCount = join("set", name, "count");
			var getArrayElement = function(index){
				return this[_array][index];
			};
			var getStaticElement = function(){
				return this[_array];
			};
			var getArray = function(){
				return this[_array].concat();
			};
			var getTempArray = function(){
				var i, a = [], max = this[_getCount]();
				for(i=0; i<max; i++) {a[i] = this[_getProperty](i)}
				return a;
			};
			this[_setProperty] = function(value, index){
				if(typeof value == "function"){
					this[_getProperty] = value; 
					this[_getArray] = getTempArray;
				}
				else if (arguments.length==1) {
					this[_array] = value;
					this[_getProperty] = getStaticElement;
					this[_getArray] = getTempArray;
				}
				else {
					if (this[_getArray] != getArray) {this[_array] = this[_getArray]()}
					this[_array][index] = value;
					this[_getProperty] = getArrayElement;
					this[_getArray] = getArray;
				}
				this[updateModel](property);
			};
			this[_setArray] = function(value){
				if(typeof value == "function"){
					this[_getArray] = value;
				}
				else {
					this[_array] = value.concat();
					this[_getProperty] = getArrayElement;
					this[_getArray] = getArray;
					this[_setCount](value.length);
				}
				this[updateModel](property);
			};
			this[_setProperty](defaultValue);
		};
//		------------------------------------------------------------
		var proxyPrototype = new Active.System.Model;
		proxyPrototype.getProperty = function(property, a, b, c){
			return this._target[getProperty](property, a, b, c);
		};
		proxyPrototype.setProperty = function(property, value, a, b, c){
			return this._target[setProperty](property, value, a, b, c);
		};
		var proxy = join("_", name, "proxy");
		this[getModel] = function(){
			if (this[external]) {return this[external]}
			if (!this[proxy]) {
				this[proxy] = proxyPrototype.clone();
				this[proxy]._target = this;
				this[proxy].$owner = this.$owner; 
			}
			return this[proxy];
		};
		_super[setModel] = function(model){
			this[external] = model;
			if (model && !model.$owner) {model.$owner = this}
		};
		_super[getModel] = function(a, b, c){
			if (this[external]) {return this[external]}
			return this.$owner[getModel](a, b, c);
		};
//		------------------------------------------------------------
		this[updateModel] = function(){};
	};
/****************************************************************
	Creates a new property for the built-in data model.
	@param	name	(String) Name of the property.
	@param	value	(Any) Default value for the property.
*****************************************************************/
	obj.defineProperty = function(name, defaultValue){
		if (name.match(_pattern)) {
			var defineProperty = join("define", RegExp.$1, "property");
			if (this[defineProperty]) {return this[defineProperty](RegExp.$2, defaultValue)}
		}
	};
/****************************************************************
	Creates a new property array for the built-in data model.
	@param	name	(String) Name of the property.
	@param	value	(Any) Default value for the property.
*****************************************************************/
	obj.definePropertyArray = function(name, defaultValue){
		if (name.match(_pattern)) {
			var defineArray = join("define", RegExp.$1, "property", "array");
			if (this[defineArray]) {return this[defineArray](RegExp.$2, defaultValue)}
		}
	};
};
Active.System.Control.create();
/******************** Formats.String.js   ********************/
Active.Formats.String = Active.System.Format.subclass();
Active.Formats.String.create = function(){
/****************************************************************
	String data format.
*****************************************************************/
	var obj = this.prototype;
/****************************************************************
	Transforms the wire data into the primitive value.
	@param	data	(String) Wire data.
	@return		Primitive value.
*****************************************************************/
	obj.dataToValue = function(data){
		return data.toUpperCase();
	};
/****************************************************************
	Transforms the wire data into the readable text.
	@param	data	(String) Wire data.
	@return		Readable text.
*****************************************************************/
	obj.dataToText = function(data){
		return data;
	};
};
Active.Formats.String.create();
/******************** Formats.Number.js   ********************/
Active.Formats.Number = Active.System.Format.subclass();
Active.Formats.Number.create = function(){
/****************************************************************
	Number formatting class.
*****************************************************************/
	var obj = this.prototype;
/****************************************************************
	Transforms the wire data into the numeric value.
	@param	data	(String) Wire data.
	@return		Numeric value.
*****************************************************************/
	obj.dataToValue = function(data){
		return Number(data);
	};
	var noFormat = function(value){
		return "" + value;
	};
	var doFormat = function(value){
		var multiplier = this._multiplier;
		var abs = (value<0) ? -value : value;
		var delta = (value<0) ? -0.5 : +0.5;
		var rounded = (Math.round(value * multiplier) + delta)/multiplier + "";
		if (abs<1000) {return rounded.replace(this.p1, this.r1)}
		if (abs<1000000) {return rounded.replace(this.p2, this.r2)}
		if (abs<1000000000) {return rounded.replace(this.p3, this.r3)}
		return rounded.replace(this.p4, this.r4);
	};
/****************************************************************
	Allows to specify the format for the text output.
	@param	format	(String) Format pattern.
*****************************************************************/
	obj.setTextFormat = function(format){
		var pattern = /^([^0#]*)([0#]*)([ .,]?)([0#]|[0#]{3})([.,])([0#]*)([^0#]*)$/;
		var f = format.match(pattern);
		if (!f) {
			this.valueToText = noFormat;
			return;
		}
		this.valueToText = doFormat;
		var rs = f[1]; // result start
		var rg = f[3]; // result group separator;
		var rd = f[5]; // result decimal separator;
		var re = f[7]; // result end
		var decimals = f[6].length;
		this._multiplier = Math.pow(10, decimals);
		var ps = "^(-?\\d+)", pm = "(\\d{3})", pe = "\\.(\\d{" + decimals + "})\\d$";
		this.p1 = new RegExp(ps + pe);
		this.p2 = new RegExp(ps + pm + pe);
		this.p3 = new RegExp(ps + pm + pm + pe);
		this.p4 = new RegExp(ps + pm + pm + pm + pe);
		this.r1 = rs + "$1" + rd + "$2" + re;
		this.r2 = rs + "$1" + rg + "$2" + rd + "$3" + re;
		this.r3 = rs + "$1" + rg + "$2" + rg + "$3" + rd + "$4" + re;
		this.r4 = rs + "$1" + rg + "$2" + rg + "$3" + rg + "$4" + rd + "$5" + re;
	};
	obj.setTextFormat("#.##");
};
Active.Formats.Number.create();
/******************** Formats.Date.js     ********************/
Active.Formats.Date = Active.System.Format.subclass();
Active.Formats.Date.create = function(){
/****************************************************************
	Date formatting class.
*****************************************************************/
	var obj = this.prototype;
	obj.date = new Date();
	obj.digits = [];
	obj.shortMonths = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
	obj.longMonths = ["January","February","March","April","May","June","July","August","September","October","November","December"];
	obj.shortWeekdays = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
	obj.longWeekdays = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
	for(var i=0; i<100; i++){obj.digits[i] = i<10 ? "0" + i : "" + i}
	var tokens = {
		"hh"	: "this.digits[this.date.getUTCHours()]",
		":mm"	: "':'+this.digits[this.date.getUTCMinutes()]",
		"mm:"	: "this.digits[this.date.getUTCMinutes()]+':'",
		"ss"	: "this.digits[this.date.getUTCSeconds()]",
		"dddd"	: "this.longWeekdays[this.date.getUTCDay()]",
		"ddd"	: "this.shortWeekdays[this.date.getUTCDay()]",
		"dd"	: "this.digits[this.date.getUTCDate()]",
		"d"		: "this.date.getUTCDate()",
		"mmmm"	: "this.longMonths[this.date.getUTCMonth()]",
		"mmm"	: "this.shortMonths[this.date.getUTCMonth()]",
		"mm"	: "this.digits[this.date.getUTCMonth()+1]",
		"m"		: "(this.date.getUTCMonth()+1)",
		"yyyy"	: "this.date.getUTCFullYear()",
		"yy"    : "this.digits[this.date.getUTCFullYear()%100]" };
	var match = "";
	for(i in tokens){
		if (typeof(i) == "string"){
			match += "|" + i;
		}
	}
	var re = new RegExp(match.replace("|", "(")+")", "gi");
/****************************************************************
	Allows to specify the format for the text output.
	@param	format	(String) Format pattern.
*****************************************************************/
	obj.setTextFormat = function(format){
		format = format.replace(re, function(i){return "'+" + tokens[i.toLowerCase()] + "+'"});
		format = "if (isNaN(value) || (value === this._valueError)) return this._textError;" +
				 "this.date.setTime(value + this._textTimezoneOffset);" +
				("return '" + format + "'").replace(/(''\+|\+'')/g, "");
		this.valueToText = new Function("value", format);
	};
	var xmlExpr = /^(....).(..).(..).(..).(..).(..)........(...).(..)/;
	var xmlOut = "$1/$2/$3 $4:$5:$6 GMT$7$8";
	var auto = function(data){
		var value = Date.parse(data + this._dataTimezoneCode);
		return isNaN(value) ? this._valueError : value;
	};
	var RFC822 = function(data){
		var value = Date.parse(data);
		return isNaN(value) ? this._valueError : value;
	};
	var ISO8061 = function(data){
		var value = Date.parse(data.replace(xmlExpr, xmlOut));
		return isNaN(value) ? this._valueError : value;
	};
/****************************************************************
	Allows to specify the wire format for data input.
	@param	format	(String) Format pattern.
*****************************************************************/
	obj.setDataFormat = function(format){
		if (format == "RFC822") {
			this.dataToValue = RFC822;
		}
		else if (format == "ISO8061") {
			this.dataToValue = ISO8061;
		}
		else {
			this.dataToValue = auto;
		}
	};
/****************************************************************
	Allows to specify the timezone used for the text output.
	@param	value	(Number) Timezone offset.
*****************************************************************/
	obj.setTextTimezone = function(value){
		this._textTimezoneOffset = value;
	};
/****************************************************************
	Allows to specify the timezone used for the data input.
	@param	value	(Number) Timezone offset.
*****************************************************************/
	obj.setDataTimezone = function(value){
		if (!value) {
			this._dataTimezoneCode = " GMT";
		}
		else {
			this._dataTimezoneCode = " GMT" +
				(value>0 ? "+" : "-") +
				this.digits[Math.floor(Math.abs(value/3600000))] +
				this.digits[Math.abs(value/60000)%60];
		}
	};
	var localTimezone = - obj.date.getTimezoneOffset() * 60000;
	obj.setTextTimezone(localTimezone);
	obj.setDataTimezone(localTimezone);
	obj.setTextFormat("d mmm yy");
	obj.setDataFormat("default");
};
Active.Formats.Date.create();
/******************** Html.Tags.js        ********************/
Active.HTML.define = function(name, tag, type){
	if (!tag) {tag = name.toLowerCase()}
	Active.HTML[name] = Active.System.HTML.subclass();
	Active.HTML[name].create = function(){};
	Active.HTML[name].prototype.setTag(tag);
};
//	------------------------------------------------------------
Active.HTML.define("DIV");
Active.HTML.define("SPAN");
Active.HTML.define("IMG");
Active.HTML.define("INPUT");
Active.HTML.define("BUTTON");
Active.HTML.define("TEXTAREA");
Active.HTML.define("TABLE");
Active.HTML.define("TR");
Active.HTML.define("TD");
/******************** Templates.Status.js ********************/
Active.Templates.Status = Active.System.Template.subclass();
Active.Templates.Status.create = function(){
/****************************************************************
	Displays status text.
*****************************************************************/
	var obj = this.prototype;
	obj.setClass("templates", "status");
	var image = new Active.HTML.SPAN;
	image.setClass("box", "image");
	image.setClass("image", function(){return this.getStatusProperty("image")});
	obj.setContent("image", image);
	obj.setContent("text", function(){return this.getStatusProperty("text")});
};
Active.Templates.Status.create();
/******************** Templates.Error.js  ********************/
Active.Templates.Error = Active.System.Template.subclass();
Active.Templates.Error.create = function(){
/****************************************************************
	Displays error information.
*****************************************************************/
	var obj = this.prototype;
	obj.setClass("templates", "error");
	obj.setContent("title", "Error: ");
	obj.setContent("text", function(){return this.getErrorProperty("text")});
};
Active.Templates.Error.create();
/******************** Templates.Text.js   ********************/
Active.Templates.Text = Active.System.Template.subclass();
Active.Templates.Text.create = function(){
/****************************************************************
	Simple text template.
*****************************************************************/
	var obj = this.prototype;
	obj.setClass("templates", "text");
	obj.setContent("text", function(){return this.getItemProperty("text")});
	obj.setEvent("onclick", function(){this.action("click")});
	obj.setEvent("ondblclick", function(){this.action("dblclick")});
};
Active.Templates.Text.create();
/******************** Templates.Image.js  ********************/
Active.Templates.Image = Active.System.Template.subclass();
Active.Templates.Image.create = function(){
/****************************************************************
	Image template.
*****************************************************************/
	var obj = this.prototype;
	obj.setClass("templates", "image");
	var image = new Active.HTML.SPAN;
	image.setClass("box", "image");
	image.setClass("image", function(){return this.getItemProperty("image")});
	obj.setContent("image", image);
	obj.setContent("text", function(){return this.getItemProperty("text")});
	obj.setEvent("onclick", function(){this.action("click")});
	obj.setEvent("ondblclick", function(){this.action("dblclick")});
};
Active.Templates.Image.create();
/******************** Templates.Link.js   ********************/
Active.Templates.Link = Active.System.Template.subclass();
Active.Templates.Link.create = function(){
/****************************************************************
	Hyperlink template.
*****************************************************************/
	var obj = this.prototype;
	obj.setTag("a");
	obj.setClass("templates", "link");
	obj.setAttribute("href", function(){return this.getItemProperty("link")});
	var image = new Active.HTML.SPAN;
	image.setClass("box", "image");
	image.setClass("image", function(){return this.getItemProperty("image")});
	obj.setContent("image", image);
	obj.setContent("text", function(){return this.getItemProperty("text")});
	obj.setEvent("onclick", function(){this.action("click")});
	obj.setEvent("ondblclick", function(){this.action("dblclick")});
};
Active.Templates.Link.create();
/******************** Templates.Item.js   ********************/
Active.Templates.Item = Active.System.Template.subclass();
Active.Templates.Item.create = function(){
/****************************************************************
	List item template.
*****************************************************************/
	var obj = this.prototype;
//	------------------------------------------------------------
	obj.setClass("templates", "item");
	obj.setClass("box", "normal");
//	------------------------------------------------------------
	var box = new Active.HTML.DIV;
	var image = new Active.HTML.SPAN;
	box.setClass("box", "item");
	image.setClass("box", "image");
	image.setClass("image", function(){return this.getItemProperty("image")});
	obj.setContent("box", box);
	obj.setContent("box/image", image);
	obj.setContent("box/text", function(){return this.getItemProperty("text")});
//	------------------------------------------------------------
//	obj.setEvent("onclick", function(){this.action("click")});
//	------------------------------------------------------------
//	obj.setEvent("onmouseenter", "mouseover(this, 'active-item-over')");
//	obj.setEvent("onmouseleave", "mouseout(this, 'active-item-over')");
};
Active.Templates.Item.create();
/******************** Templates.List.js   ********************/
Active.Templates.List = Active.System.Template.subclass();
Active.Templates.List.create = function(){
/****************************************************************
	List box template.
*****************************************************************/
	var obj = this.prototype;
//	list does not have html element (provides content only)
	obj.setTag("");
	obj.defineTemplate("item", new Active.Templates.Text);
//	redirect item property request to data property (index)
	var getItemProperty = function(property){
		return this.$owner.getDataProperty(property, this.$index);
	};
	var setItemProperty = function(property, value){
		return this.$owner.setDataProperty(property, value, this.$index);
	};
	obj.getItemTemplate = function(index, temp){
		var template = this.defaultItemTemplate(index);
		if (!temp) {temp = []}
		if (!temp.selected) {
			temp.selected = [];
			var i, values = this.getSelectionProperty("values");
			for (i=0; i<values.length; i++) {temp.selected[values[i]]=true}
			template.getItemProperty = getItemProperty;
			template.setItemProperty = setItemProperty;
			template.setClass("list", "item");
		}
		if (temp.selected[index]){
			template = template.clone();
			template.$index = "";
			template.setClass("selection", true);
			template.$index = index;
		}
		return template;
	};
//	------------------------------------------------------------
	var html = function(){
		var i, result = [], temp = [], items = this.getItemsProperty("values");
		for(i=0; i<items.length; i++) {result[i] = this.getItemTemplate(items[i], temp).toString()}
		return result.join("");
	};
	obj.setContent("html", html);
//	------------------------------------------------------------
};
Active.Templates.List.create();
/******************** Templates.Row.js    ********************/
Active.Templates.Row = Active.Templates.List.subclass();
Active.Templates.Row.create = function(){
/****************************************************************
	Grid row template.
*****************************************************************/
	var obj = this.prototype;
	var _super = this.superclass.prototype;
//	------------------------------------------------------------
	obj.setTag("div");
	obj.setClass("templates", "row");
	obj.setClass("grid", "row");
//	------------------------------------------------------------
	obj.getDataProperty = function(property, i){
		return this.$owner.getDataProperty(property, this.$index, i);
	};
	obj.setDataProperty = function(property, value, i){
		return this.$owner.setDataProperty(property, value, this.$index, i);
	};
	obj.getItemsProperty = function(property){
		return this.getColumnProperty(property);
	};
	obj.getSelectionProperty = function(property){
		return this.getDummyProperty(property);
	};
	obj.getRowProperty = function(property){
		return this.$owner.getItemsProperty(property, this.$index);
	};
//	------------------------------------------------------------
	var getItemProperty = function(property){
		return this.$owner.getDataProperty(property, this.$index);
	};
	var setItemProperty = function(property, value){
		return this.$owner.setDataProperty(property, value, this.$index);
	};
	var getColumnProperty = function(property){
		return this.$owner.getColumnProperty(property, this.$index);
	};
	obj.getItemTemplate = function(i){
		if (!this._itemTemplates) {
			this._itemTemplates = [];
		}
		if (this._itemTemplates[i]) {
			this._itemTemplates[i]._id = this._id + ".item:" + i;
			this._itemTemplates[i].$owner = this; // 1.0.1/01 selected first row
			return this._itemTemplates[i];
		}
		if (typeof(i)=="undefined") {return _super.getItemTemplate.call(this)}
		var template = _super.getItemTemplate.call(this, i).clone();
		template.$index = i;
		template.setClass("column", i);
		this._itemTemplates[i] = template;
		return template;
	};
	obj.setItemTemplate = function(template, i){
		template.getItemProperty = getItemProperty;
		template.setItemProperty = setItemProperty;
		template.getColumnProperty = getColumnProperty;
		template.setClass("row", "cell");
		template.setClass("grid", "column");
		if (typeof(i)=="undefined") {return _super.setItemTemplate.call(this, template)}
		template.setClass("column", i);
		template.$owner = this;
		template.$index = i;
		if (!this._itemTemplates) {
			this._itemTemplates = [];
		}
		this._itemTemplates[i] = template;
	};
//	------------------------------------------------------------
	var selectRow = function(event){
		if (event.shiftKey) {return this.action("selectRangeOfRows")}
		if (event.ctrlKey) {return this.action("selectMultipleRows")}
		this.action("selectRow");
	};
	obj.setEvent("onclick", selectRow);
};
Active.Templates.Row.create();
/******************** Templates.Header.js ********************/
Active.Templates.Header = Active.Templates.Item.subclass();
Active.Templates.Header.create = function(){
/****************************************************************
	Column header template.
*****************************************************************/
	var obj = this.prototype;
//	------------------------------------------------------------
	obj.setClass("templates", "header");
	obj.setClass("column", function(){return this.$index});
	obj.setClass("sort", function(){
		return this.getSortProperty("index") != this.$index ? "none" : this.getSortProperty("direction");
	});
	obj.setAttribute("title", function(){return this.getItemProperty("tooltip")});
//	------------------------------------------------------------
	var div = new Active.HTML.DIV;
	div.setClass("box", "resize");
	div.setEvent("onmousedown", function(){this.action("startColumnResize")});
	div.setContent("html", "&nbsp;"); 

	obj.setContent("div", div);
	obj.setEvent("onmousedown", function(){
		this.setClass("header", "pressed");
		window.status = "Sorting...";
		this.timeout(function(){this.action("columnSort")});
	});
	var sort = new Active.HTML.SPAN;
	sort.setClass("box", "sort");
	obj.setContent("box/sort", sort);
	obj.setEvent("onmouseenter", "mouseover(this, 'active-header-over')");
	obj.setEvent("onmouseleave", "mouseout(this, 'active-header-over')");
};
Active.Templates.Header.create();
/******************** Templates.Box.js    ********************/
Active.Templates.Box = Active.System.Template.subclass();
Active.Templates.Box.create = function(){
/****************************************************************
	Generic 'box' template.
*****************************************************************/
	var obj = this.prototype;
//	------------------------------------------------------------
	obj.setClass("templates", "box");
	obj.setClass("box", "normal");
//	------------------------------------------------------------
	var box = new Active.HTML.DIV;
	box.setClass("box", "item");
	obj.setContent("box", box);
//	------------------------------------------------------------
};
Active.Templates.Box.create();
/******************** Templates.Scroll.js ********************/
Active.Templates.Scroll = Active.System.Template.subclass();
Active.Templates.Scroll.create = function(){
/****************************************************************
	Four panes scrollable layout template.
*****************************************************************/
	var obj = this.prototype;
	var _super = this.superclass.prototype;
//	------------------------------------------------------------
	obj.setTag("");
//	------------------------------------------------------------
	var Pane = Active.HTML.DIV;
	var Box = Active.Templates.Box;
	var data = new Pane;
	var top = new Pane;
	var left = new Pane;
	var corner = new Box;
	var fill = new Box;
	var scrollbars = new Pane;
	var space = new Pane;
	data.setClass("scroll", "data");
	top.setClass("scroll", "top");
	left.setClass("scroll", "left");
	corner.setClass("scroll", "corner");
	fill.setClass("scroll", "fill");
	scrollbars.setClass("scroll", "bars");
	space.setClass("scroll", "space");
	obj.setContent("data", data);
	obj.setContent("top", top);
	obj.setContent("left", left);
	obj.setContent("corner", corner);
	obj.setContent("scrollbars", scrollbars);
	obj.setContent("data/html", function(){return this.getMainTemplate()});
	obj.setContent("top/html", function(){return this.getTopTemplate()});
	obj.setContent("left/html", function(){return this.getLeftTemplate()});
	obj.setContent("scrollbars/space", space);
	obj.setContent("top/fill", fill);
//	------------------------------------------------------------
	var scroll = function(){
		var scrollbars = this.getContent("scrollbars").element();
		var data = this.getContent("data").element();
		var top = this.getContent("top").element();
		var left = this.getContent("left").element();
		var x = scrollbars.scrollLeft;
		var y = scrollbars.scrollTop;
		data.scrollLeft = x;
		top.scrollLeft = x;
		data.scrollTop = y;
		left.scrollTop = y;
		scrollbars = null;
		data = null;
		top = null;
		left = null;
	};
	scrollbars.setEvent("onscroll", scroll);
//	------------------------------------------------------------
	var resize = function(){
		if (this._sizeAdjusted){
			this._sizeAdjusted = false;
			this.timeout(adjustSize, 100);
			var data = this.getContent("data").element();
			var scrollbars = this.getContent("scrollbars").element();
			var top = this.getContent("top").element();
			var left = this.getContent("left").element();
			data.runtimeStyle.width = "100%";
			top.runtimeStyle.width = "100%";
			data.runtimeStyle.height = "100%";
			left.runtimeStyle.height = "100%";
			scrollbars.runtimeStyle.zIndex = 1000;
			data = null;
			scrollbars = null;
			top = null;
			left = null;
		}
	};
	scrollbars.setEvent("onresize", resize);
//	------------------------------------------------------------
	obj._sizeAdjusted = true;
	var adjustSize = function(){
		var data = this.getContent("data").element();
		var scrollbars = this.getContent("scrollbars").element();
		var top = this.getContent("top").element();
		var left = this.getContent("left").element();
		var space = this.getContent("scrollbars/space").element();
		if (data) {
			if (data.scrollHeight) {
				space.runtimeStyle.height = data.scrollHeight > data.offsetHeight ? data.scrollHeight : 0;
				space.runtimeStyle.width = data.scrollWidth > data.offsetWidth ? data.scrollWidth : 0;
				var y = scrollbars.clientHeight;
				var x = scrollbars.clientWidth;
				data.runtimeStyle.width = x;
				top.runtimeStyle.width = x;
				data.runtimeStyle.height = y;
				left.runtimeStyle.height = y;
				top.scrollLeft = data.scrollLeft;
				left.scrollTop = data.scrollTop;
				scrollbars.runtimeStyle.zIndex = 0;
			}
			else {
				this.timeout(adjustSize, 500);
			}
			data.className = data.className + "";
		}
		data = null;
		scrollbars = null;
		top = null;
		left = null;
		space = null;
		this._sizeAdjusted = true;
	};
	// delay for grid col resize
	obj.setAction("adjustSize", function(){this.timeout(adjustSize, 500);});
	obj.toString = function(){
		this.timeout(adjustSize);
		return _super.toString.call(this);
	};
};
Active.Templates.Scroll.create();
/******************** Controls.Grid.js    ********************/
Active.Controls.Grid = Active.System.Control.subclass();
Active.Controls.Grid.create = function(){
/****************************************************************
	Scrollable grid control. Displays data in a table with fixed
	headers, resizable columns, client-side sorting and much more.
*****************************************************************/
	var obj = this.prototype;
	obj.setClass("controls", "grid");
	obj.setAttribute("tabIndex", "-1");
	obj.setAttribute("hideFocus", "true");
/****************************************************************
	Splits the grid display into the four scrolling areas.
*****************************************************************/
	obj.defineTemplate("layout", new Active.Templates.Scroll);
/****************************************************************
	Contains the main area of the grid.
*****************************************************************/
	obj.defineTemplate("main", function(){
		switch (this.getStatusProperty("code")) {
			case "":
				return this.getDataTemplate();
			case "error":
				return this.getErrorTemplate();
			default:
				return this.getStatusTemplate();
		}
	});
/****************************************************************
	Contains the list of data rows.
*****************************************************************/
	obj.defineTemplate("data", new Active.Templates.List);
/****************************************************************
	Contains the row headings area.
*****************************************************************/
	obj.defineTemplate("left", new Active.Templates.List);
/****************************************************************
	Contains the column headings area.
*****************************************************************/
	obj.defineTemplate("top", new Active.Templates.List);
/****************************************************************
	Displays control status text.
*****************************************************************/
	obj.defineTemplate("status", new Active.Templates.Status);
/****************************************************************
	Displays error text.
*****************************************************************/
	obj.defineTemplate("error",	new Active.Templates.Error);
/****************************************************************
	Grid row template.
*****************************************************************/
	obj.defineTemplate("row", new Active.System.Template);
/****************************************************************
	Grid column (cell) template.
*****************************************************************/
	obj.defineTemplate("column", new Active.System.Template);
	obj.getColumnTemplate = function(i){return this.getTemplate("data/item/item", i)};
	obj.setColumnTemplate = function(template, i){this.setTemplate("data/item/item", template, i)};
	obj.getRowTemplate = function(i){return this.getTemplate("data/item", i)};
	obj.setRowTemplate = function(template, i){this.setTemplate("data/item", template, i)};
	obj.setTemplate("data/item", 	new Active.Templates.Row);
	obj.setTemplate("left/item",	new Active.Templates.Item);
	obj.setTemplate("top/item", 	new Active.Templates.Header);
/****************************************************************
	Specifies the row indices and the row headers data.
	It defines which data rows and in which order should be requested
	from the data model for the grid display.
*****************************************************************/
	obj.defineModel("row");
/****************************************************************
	Sets or retrieves the number of rows in the grid.
	@remarks
	Setting row count will re-initialize row values array to 0..count-1
*****************************************************************/
	obj.defineRowProperty("count", function(){return this.getDataProperty("count")} );
/****************************************************************
	Retrieves the row index.
*****************************************************************/
	obj.defineRowProperty("index", function(i){return i});
/****************************************************************
	Retrieves the display order for the row.
*****************************************************************/
	obj.defineRowProperty("order", function(i){return i});
/****************************************************************
	Allows to specify the text for the row headers.
*****************************************************************/
	obj.defineRowPropertyArray("text", function(i){return this.getRowOrder(i) + 1});
/****************************************************************
	Allows to specify the image to display in the row headers.
*****************************************************************/
	obj.defineRowPropertyArray("image", "none");
/****************************************************************
	Sets or retrieves the row index or the array of indexes.
*****************************************************************/
	obj.defineRowPropertyArray("value", function(i){return i});
/****************************************************************
	Specifies the column indices and the column headers data.
	Defines which data items should be displayed in each column.
*****************************************************************/
	obj.defineModel("column");
/****************************************************************
	Sets or retrieves the number of columns in the grid.
*****************************************************************/
	obj.defineColumnProperty("count", 0 );
/****************************************************************
	Retrieves the column index.
*****************************************************************/
	obj.defineColumnProperty("index", function(i){return i});
/****************************************************************
	Retrieves the display order for the column.
*****************************************************************/
	obj.defineColumnProperty("order", function(i){return i});
/****************************************************************
	Allows to specify the text for the column headers.
*****************************************************************/
	obj.defineColumnPropertyArray("text", function(i){return "Column " + i});
/****************************************************************
	Allows to specify the image to display in the column headers.
*****************************************************************/
	obj.defineColumnPropertyArray("image", "none");
/****************************************************************
	Sets or retrieves the column index or the array of indexes.
*****************************************************************/
	obj.defineColumnPropertyArray("value", function(i){return i});
/****************************************************************
	Allows to specify the tooltips text for the column headers.
*****************************************************************/
	obj.defineColumnPropertyArray("tooltip", "");
/****************************************************************
	Provides the content to display inside the grid cells.
*****************************************************************/
	obj.defineModel("data");
/****************************************************************
	Sets or retrieves the number of data items (rows).
*****************************************************************/
	obj.defineDataProperty("count", 0);
/****************************************************************
	Retrieves the data item index (row).
*****************************************************************/
	obj.defineDataProperty("index", function(i){return i});
/****************************************************************
	Allows to specify the text for the grid cells.
*****************************************************************/
	obj.defineDataProperty("text", "");
/****************************************************************
	Allows to specify the image to display in the grid cells.
*****************************************************************/
	obj.defineDataProperty("image", "none");
/****************************************************************
	Allows to specify the link URL for a cell.
	Use Active.Templates.Link as a column template.
*****************************************************************/
	obj.defineDataProperty("link", "");
/****************************************************************
	Provides the value to be used for sorting the data.
*****************************************************************/
	obj.defineDataProperty("value", function(i,j){
		var text = "" + this.getDataText(i, j);
		var value = Number(text.replace(/[ ,%\$]/gi, "").replace(/\((.*)\)/, "-$1"));
		return isNaN(value) ? text.toLowerCase() + " " : value;
	});
/****************************************************************
	Items model.
*****************************************************************/
	obj.defineModel("items"); 

/****************************************************************
	Used as a stub where no actual data is required.
*****************************************************************/
	obj.defineModel("dummy");
	obj.defineDummyProperty("count", 0);
	obj.defineDummyPropertyArray("value", -1);
/****************************************************************
	Controls the row/column/cell selection.
*****************************************************************/
	obj.defineModel("selection");
/****************************************************************
	Sets or retrieves the active cell index.
*****************************************************************/
	obj.defineSelectionProperty("index", -1); 

/****************************************************************
	Specifies if multiple selection is allowed.
*****************************************************************/
	obj.defineSelectionProperty("multiple", false);
/****************************************************************
	Provides the number of selected items.
*****************************************************************/
	obj.defineSelectionProperty("count", 0);
/****************************************************************
	Provides the array of the selected item indices.
*****************************************************************/
	obj.defineSelectionPropertyArray("value", 0);
/****************************************************************
	Controls sorting of the grid rows.
*****************************************************************/
	obj.defineModel("sort");
/****************************************************************
	Specifies the index of a column to sort data on.
*****************************************************************/
	obj.defineSortProperty("index", -1);
/****************************************************************
	Specifies the sort direction.
*****************************************************************/
	obj.defineSortProperty("direction", "none");
/****************************************************************
	Provides control status.
*****************************************************************/
	obj.defineModel("status");
/****************************************************************
	Provides status code.
*****************************************************************/
	obj.defineStatusProperty("code", function(){
		var data = this.getDataModel();
		if (!data.isReady()) {
			return "loading";
		}
		if (!this.getRowProperty("count")) {
			return "nodata";
		}
		return "";
	});
/****************************************************************
	Provides status text.
*****************************************************************/
	obj.defineStatusProperty("text", function(){
		switch(this.getStatusProperty("code")) {
			case "loading":
				return "Loading data, please wait...";
			case "nodata":
				return "No record found.";
			default:
				return "";
		}
	});
/****************************************************************
	Provides status image.
*****************************************************************/
	obj.defineStatusProperty("image", function(){
		switch(this.getStatusProperty("code")) {
			case "loading":
				return "loading";
			default:
				return "none";
		}
	});
/****************************************************************
	Provides error information.
*****************************************************************/
	obj.defineModel("error");
/****************************************************************
	Provides error code.
*****************************************************************/
	obj.defineErrorProperty("code", 0);
/****************************************************************
	Provides error text.
*****************************************************************/
	obj.defineErrorProperty("text", "");
//	------------------------------------------------------------
//	------------------------------------------------------------
	obj.getLeftTemplate = function(){
		var template = this.defaultLeftTemplate();
		template.setDataModel(this.getRowModel());
		template.setItemsModel(this.getRowModel());
		template.setSelectionModel(this.getDummyModel());
		return template;
	};
//	------------------------------------------------------------
	obj.getTopTemplate = function(){
		var template = this.defaultTopTemplate();
		template.setDataModel(this.getColumnModel());
		template.setItemsModel(this.getColumnModel());
		template.setSelectionModel(this.getDummyModel());
		return template;
	};
//	------------------------------------------------------------
	obj.getDataTemplate = function(){
		var template = this.defaultDataTemplate();
		template.setDataModel(this.getDataModel());
		template.setItemsModel(this.getRowModel());
		return template;
	};
//	------------------------------------------------------------
	obj.setContent(function(){return this.getLayoutTemplate()});
/****************************************************************
	Allows to specify the height of the column headers.
	@param	height (Number) The new height value.
*****************************************************************/
	obj.setColumnHeaderHeight = function(height){
		var layout = this.getTemplate("layout");
		layout.getContent("top").setStyle("height", height);
		layout.getContent("corner").setStyle("height", height);
		layout.getContent("left").setStyle("padding-top", height);
		layout.getContent("data").setStyle("padding-top", height);
	};
/****************************************************************
	Allows to specify the width of the row headers.
	@param	width (Number) The new width value.
*****************************************************************/
	obj.setRowHeaderWidth = function(width){
		var layout = this.getTemplate("layout");
		layout.getContent("left").setStyle("width", width);
		layout.getContent("corner").setStyle("width", width);
		layout.getContent("top").setStyle("padding-left", width);
		layout.getContent("data").setStyle("padding-left", width);
	};
//	------------------------------------------------------------
	var startColumnResize = function(header){
		
		var el = header.element();
		var pos = event.clientX;
		var size = el.offsetWidth;
		var grid = this;
		var doResize = function(){
			var el = header.element();
			var sz = size + event.clientX - pos;
			el.style.width = sz < 10 ? 10 : sz;
			el = null;
		};
		var endResize = function(){
			var el = header.element();
			if( typeof el.onmouseleave == "function") {
				el.onmouseleave();
			}
			el.detachEvent("onmousemove", doResize);
			el.detachEvent("onmouseup", endResize);
			el.detachEvent("onlosecapture", endResize);
			el.releaseCapture();
			var width = size + event.clientX - pos;
			if (width < 10) {width = 10}
			el.style.width = width;
			var ss = document.styleSheets[document.styleSheets.length-1];
			var i, selector = "#" + grid.getId() + " .active-column-" + header.getItemProperty("index");
			for(i=0; i<ss.rules.length;i++){
				if(ss.rules[i].selectorText == selector){
					ss.rules[i].style.width = width;
					el = null;
					grid.getTemplate("layout").action("adjustSize");
					return; 
				}
			}
			ss.addRule(selector, "width:" + width + "px");
			el = null;
			grid.getTemplate("layout").action("adjustSize");
		};
		el.attachEvent("onmousemove", doResize);
		el.attachEvent("onmouseup", endResize);
		el.attachEvent("onlosecapture", endResize);
		el.setCapture();
//		break object reference to avoid memory leak
		el = null;
		event.cancelBubble = true;
	};
	obj.setAction("startColumnResize", startColumnResize);
//	------------------------------------------------------------
	var setSelectionIndex = obj.setSelectionIndex;
	obj.setSelectionIndex = function(index){
		setSelectionIndex.call(this, index);
		this.setSelectionValues([index]);
		var row = this.getTemplate("row", index);
		var data = this.getTemplate("layout").getContent("data");
		var left = this.getTemplate("layout").getContent("left");
		var scrollbars = this.getTemplate("layout").getContent("scrollbars");
		try {
			var top, padding = data.element().firstChild.offsetTop; // 1.0.2 - allows set header height in em instead of px
			if (data.element().scrollTop > row.element().offsetTop - padding) {
				top = row.element().offsetTop  - padding;
				left.element().scrollTop = top;
				data.element().scrollTop = top;
				scrollbars.element().scrollTop = top;
			}
			if (data.element().offsetHeight + data.element().scrollTop <
				row.element().offsetTop + row.element().offsetHeight ) {
				top = row.element().offsetTop + row.element().offsetHeight - data.element().offsetHeight;
				left.element().scrollTop = top;
				data.element().scrollTop = top;
				scrollbars.element().scrollTop = top;
			}
		}
		catch(error){
			// ignore errors
		}
	};
//	------------------------------------------------------------
	var setSelectionValues = obj.setSelectionValues;
	obj.setSelectionValues = function(array){
		var i, current = this.getSelectionValues();
		setSelectionValues.call(this, array);
		var changes = {};
		for (i=0; i<current.length; i++) {
			changes[current[i]] = true;
		}
		for (i=0; i<array.length; i++) {
 			changes[array[i]] = changes[array[i]] ? false : true;
		}
		for (i in changes) {
			if (changes[i]===true){
				this.getRowTemplate(i).refreshClasses();
			}
		}
		this.action("selectionChanged");
	};
//	------------------------------------------------------------
	var selectRow = function(src){
		this.setSelectionProperty("index", src.getItemProperty("index"));
	};
	var selectMultipleRows = function(src){
		if (!this.getSelectionProperty("multiple")){
			return this.action("selectRow", src);
		}
		var index = src.getItemProperty("index");
		var selection = this.getSelectionProperty("values");
		for (var i=0; i<selection.length; i++){
			if(selection[i]==index){
				selection.splice(i, 1);
				i = -1;
				break;
			}
		}
		if (i!=-1) {
			selection.push(index);
		}
		this.setSelectionProperty("values", selection);
		setSelectionIndex.call(this, index);
		this.getRowTemplate(index).refreshClasses();
		this.action("selectionChanged");
	};
	var selectRangeOfRows = function(src){
		if (!this.getSelectionProperty("multiple")){
			return this.action("selectRow", src);
		}
		var previous = this.getSelectionProperty("index");
		var index = src.getItemProperty("index");
		var row1 = Number(this.getRowProperty("order", previous));
		var row2 = Number(this.getRowProperty("order", index));
		var start = row1 > row2 ? row2 : row1;
		var count = row1 > row2 ? row1 - row2 : row2 - row1;
		var i, selection = [];
		for(i=0; i<=count; i++){
			selection.push(this.getRowProperty("value", start + i));
		}
		this.setSelectionProperty("values", selection);
		setSelectionIndex.call(this, index);
		this.getRowTemplate(index).refreshClasses();
		this.action("selectionChanged");
	};
	obj.setAction("selectRow", selectRow);
	obj.setAction("selectMultipleRows", selectMultipleRows);
	obj.setAction("selectRangeOfRows", selectRangeOfRows);
/****************************************************************
	Sorts the rows with the data in the given column.
	@param 	index (Number) Column index to sort on.
	@param 	direction (String) Sort direction ("ascending" or "descending").
*****************************************************************/
	obj.sort = function(index, direction){
		var model = this.getModel("row");
		if (model.sort) {
			return model.sort(index, direction);
		}
		function compare(value, pos, dir){
			var greater = 1, less = -1;
			if (dir == "descending"){
				greater = -1;
				less = 1;
			}
			var types = {
				"undefined"	: 0,
				"boolean"	: 1,
				"number"	: 2,
				"string"	: 3,
				"object"	: 4,
				"function"	: 5
			};
			return function(i, j){
				var a = value[i], b = value[j], x, y;
				if (typeof(a) != typeof(b)){
					x = types[typeof(a)];
					y = types[typeof(b)];
					if (x > y) {return greater}
					if (x < y) {return less}
				}
				else if (typeof(a)=="number"){
					if (a > b) {return greater}
					if (a < b) {return less}
				}
				else {
					var result = ("" + a).localeCompare(b);
					if (result) {return greater * result}
				}
				x = pos[i];
				y = pos[j];
				if (x > y) {return 1}
				if (x < y) {return -1}
				return 0;
			}
		}
		if (direction && direction != "ascending" ) {
 			direction = "descending";
 		}
 		else {
 			direction = "ascending";
 		}
		var i, value = {}, pos = {};
		var rows = this.getRowProperty("values");
		for (i=0; i<rows.length; i++) {
			value[rows[i]] = this.getDataProperty("value", rows[i], index);
			pos[rows[i]] = i;
		}
		rows.sort(compare(value, pos, direction));
		this.setRowProperty("values", rows);
		this.setSortProperty("index", index);
		this.setSortProperty("direction", direction);
	};
	obj.setAction("columnSort", function(src){
		var i = src.getItemProperty("index");
		var d = (this.getSortProperty("index") == i) && (this.getSortProperty("direction")=="ascending") ? "descending" : "ascending";
		window.status = "Sorting...";
		this.sort(i, d);
		this.refresh();
		this.timeout(function(){window.status = ""});
	});
//	------------------------------------------------------------
	var _getRowOrder = function(i){
		return this._rowOrders[i];
	};
	var _setRowValues = obj.setRowValues;
	obj.setRowValues = function(values){
		_setRowValues.call(this, values);
		var i, max = values.length, orders = [];
		for(i=0; i<max; i++){
			orders[values[i]] = i;
		}
		this._rowOrders = orders;
		this.getRowOrder = _getRowOrder;
	};
//	------------------------------------------------------------
	obj._kbSelect = function(delta){
		var index = this.getSelectionProperty("index");
		var order = this.getRowProperty("order", index );
		var count = this.getRowProperty("count");
		var newOrder = Number(order) + delta;
		if (newOrder<0) {newOrder = 0}
		if (newOrder>count-1) {newOrder = count-1}
		if (delta == -100) {newOrder = 0}
		if (delta == 100) {newOrder = count-1}
		var newIndex = this.getRowProperty("value", newOrder);
		this.setSelectionProperty("index", newIndex);
	};
	obj.setAction("up", function(){this._kbSelect(-1)});
	obj.setAction("down", function(){this._kbSelect(+1)});
	obj.setAction("pageUp", function(){this._kbSelect(-10)});
	obj.setAction("pageDown", function(){this._kbSelect(+10)});
	obj.setAction("home", function(){this._kbSelect(-100)});
	obj.setAction("end", function(){this._kbSelect(+100)});
//	------------------------------------------------------------
	var kbActions = {
		38 : "up",
		40 : "down",
		33 : "pageUp",
		34 : "pageDown",
		36 : "home",
		35 : "end"	};
	var onkeydown = function(event){
		var action = kbActions[event.keyCode];
		if (action)	{
			this.action(action);
			event.returnValue = false;
			event.cancelBubble = true;
		}
	};
	obj.setEvent("onkeydown", onkeydown);
//	------------------------------------------------------------
	function onmousewheel(event){
		var scrollbars = this.getTemplate("layout").getContent("scrollbars");
		var delta = scrollbars.element().offsetHeight * event.wheelDelta/480;
		scrollbars.element().scrollTop -= delta;
		event.returnValue = false;
		event.cancelBubble = true;
	}
	obj.setEvent("onmousewheel", onmousewheel);
};
Active.Controls.Grid.create();
/******************** Http.Request.js     ********************/
Active.HTTP.Request = Active.System.Model.subclass();
Active.HTTP.Request.create = function(){
/****************************************************************
	Generic HTTP request class.
*****************************************************************/
	var obj = this.prototype;
/****************************************************************
	Sets or retrieves the remote data URL.
*****************************************************************/
	obj.defineProperty("URL");
/****************************************************************
	Indicates whether asynchronous download is permitted.
*****************************************************************/
	obj.defineProperty("async", true);
/****************************************************************
	Specifies HTTP request method.
*****************************************************************/
	obj.defineProperty("requestMethod", "GET");
/****************************************************************
	Allows to send data with the request.
*****************************************************************/
	obj.defineProperty("requestData", "");
/****************************************************************
	Returns response text.
*****************************************************************/
	obj.defineProperty("responseText", function(){return this._http ? this._http.responseText : ""});
/****************************************************************
	Returns response XML.
*****************************************************************/
	obj.defineProperty("responseXML", function(){return this._http ? this._http.responseXML : ""});
/****************************************************************
	Sets or retrieves the user name.
*****************************************************************/
	obj.defineProperty("username", null);
/****************************************************************
	Sets or retrieves the password.
*****************************************************************/
	obj.defineProperty("password", null);
/****************************************************************
	Allows to specify namespaces for use in XPath expressions.
	@param name (String) The namespace alias.
	@param value (String) The namespace URL.
*****************************************************************/
	obj.setNamespace = function(name, value){
		this._namespaces += " xmlns:" + name + "=\"" + value + "\"";
	};
	obj._namespaces = "";
/****************************************************************
	Allows to specify the request arguments/parameters.
	@param name (String) The parameter name.
	@param value (String) The parameter value.
*****************************************************************/
	obj.setParameter = function(name, value){
		this["_" + name + "Parameter"] = value;
		if ((this._parameters + " ").indexOf(" " + name + " ") < 0) {
			this._parameters += " " + name;
		}
	};
	obj._parameters = "";
/****************************************************************
	Sets HTTP request header.
	@param name (String) The request header name.
	@param value (String) The request header value.
*****************************************************************/
	obj.setRequestHeader = function(name, value){
		this["_" + name + "Header"] = value;
		if ((this._headers + " ").indexOf(" " + name + " ") < 0) {
			this._headers += " " + name;
		}
	};
	obj._headers = "";
	obj.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
/****************************************************************
	Returns HTTP response header (for example "Content-Type").
*****************************************************************/
	obj.getResponseHeader = function(name){
		return this._http ? this._http.getResponseHeader(name) : "";
	};
/****************************************************************
	Sends the request.
*****************************************************************/
	obj.request = function(){
		var self = this;
		this._ready = false;
		var i, name, value, data = "", params = this._parameters.split(" ");
		for (i=1; i<params.length; i++){
			name = params[i];
			value = this["_" + name + "Parameter"];
			if (typeof value == "function") { value = value(); }
			data += name + "=" + encodeURIComponent(value) + "&";
		}
		var URL = this._URL;
		if ((this._requestMethod != "POST") && data) {
			URL += "?" + data;
			data = null;
		}
		this._http = window.ActiveXObject ? new ActiveXObject("MSXML2.XMLHTTP") : new XMLHttpRequest;
		this._http.open(this._requestMethod, URL, this._async, this._username, this._password);
		var headers = this._headers.split(" ");
		for (i=1; i<headers.length; i++){
			name = headers[i];
			value = this["_" + name + "Header"];
			if (typeof value == "function") { value = value(); }
			this._http.setRequestHeader(name, value);
		}
		this._http.send(data);
		if (this._async) {
			this.timeout(wait, 200);
		}
		else {
			returnResult();
		}
		function wait(){
			if (self._http.readyState == 4) {
				self._ready = true;
				returnResult();
			}
			else {
				self.timeout(wait, 200);
			}
		}
		function returnResult(){
			if (self._http.responseXML && self._http.responseXML.hasChildNodes()) {
				self.response(self._http.responseXML);
			}
			else {
				self.response(self._http.responseText);
			}
		}
	};
/****************************************************************
	Allows to process the received data.
	@param result (Object) The downloaded data (XML DOMDocument object).
*****************************************************************/
	obj.response = function(result){
		if (this.$owner) {this.$owner.refresh()}
	};
/****************************************************************
	Indicates whether the request is already completed.
*****************************************************************/
	obj.isReady = function(){
		return this._ready;
	};
};
Active.HTTP.Request.create();
/******************** Text.Table.js       ********************/
Active.Text.Table = Active.HTTP.Request.subclass();
Active.Text.Table.create = function(){
/****************************************************************
	Table model for loading and parsing data in CSV text format.
*****************************************************************/
	var obj = this.prototype;
	var _super = this.superclass.prototype;
/****************************************************************
	Allows to process the received text.
	@param text (String) The downloaded text.
*****************************************************************/
	obj.response = function(text){
		var i, s, table = [], a = text.split(/\r*\n/);
		var pattern = new RegExp("(^|\\t|,)(\"*|'*)(.*?)\\2(?=,|\\t|$)", "g");
		for (i=0; i<a.length; i++) {
			s = a[i].replace(/""/g, "'");
			s = s.replace(pattern, "$3\t");
			s = s.replace(/\t$/, "");
			if (s) {table[i] = s.split(/\t/)}
		}
		this._data = table;
		_super.response.call(this);
	};
	obj._data = [];
/****************************************************************
	Returns the number of data rows.
*****************************************************************/
	obj.getCount = function(){
		return this._data.length;
	};
/****************************************************************
	Returns the index.
*****************************************************************/
	obj.getIndex = function(i){
		return i;
	};
/****************************************************************
	Returns the cell text.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getText = function(i, j){
		return this._data[i][j];
	};
/****************************************************************
	Returns the cell image.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getImage = function(){
		return "none";
	};
/****************************************************************
	Returns the cell hyperlink.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getLink = function(){
		return "";
	};
/****************************************************************
	Returns the cell value.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getValue = function(i, j){
		var text = this.getText(i, j);
		var value = Number(text.replace(/[ ,%\$]/gi, "").replace(/\((.*)\)/, "-$1"));
		return isNaN(value) ? text.toLowerCase() + " " : value;
	};
};
Active.Text.Table.create();
/******************** Xml.Table.js        ********************/
Active.XML.Table = Active.HTTP.Request.subclass();
Active.XML.Table.create = function(){
/****************************************************************
	Table model for loading and parsing data in XML format.
*****************************************************************/
	var obj = this.prototype;
	var _super = this.superclass.prototype;
/****************************************************************
	Allows to process the received data.
	@param xml (DOMDocument) The received data.
*****************************************************************/
	obj.response = function(xml){
		this.setXML(xml);
		_super.response.call(this);
	};
/****************************************************************
	Sets or retrieves the XML document (or string).
*****************************************************************/
	obj.defineProperty("XML");
	obj.setXML = function(xml){
		if (!xml.nodeType) {
			var s = "" + xml;
			if (window.ActiveXObject) {
				xml = new ActiveXObject("MSXML2.DOMDocument");
				xml.loadXML(s);
				xml.setProperty("SelectionLanguage", "XPath");
			}
			else {
				xml = (new DOMParser).parseFromString(s, "text/xml");
			}
		}
		if (this._namespaces) {xml.setProperty("SelectionNamespaces", this._namespaces);}
		this._xml = xml;
		this._data = this._xml.selectSingleNode(this._dataPath);
		this._items = this._data ? this._data.selectNodes(this._itemPath) : null;
		this._ready = true;
	};
	obj.getXML = function(){
		return this._xml;
	};
	obj._dataPath = "*";
	obj._itemPath = "*";
	obj._valuePath = "*";
	obj._valuesPath = [];
	obj._formats = [];
/****************************************************************
	Sets the XPath expressions to retrieve values for each column.
	@param array (Array) The array of XPaths expressions.
*****************************************************************/
	obj.setColumns = function(array){
		this._valuesPath = array;
	};
/****************************************************************
	Specifies the XPath expression to retrieve the set of rows.
	@param xpath (String) The xpath expression.
*****************************************************************/
	obj.setRows = function(xpath){
		this._itemPath = xpath;
	};
/****************************************************************
	Specifies the XPath expression to select the table root element.
	@param xpath (String) The xpath expression.
*****************************************************************/
	obj.setTable = function(xpath){
		this._dataPath = xpath;
	};
/****************************************************************
	Allows to specify the formatting object for the column.
	@param format (Object) The formatting object.
	@param index (Index) The column index.
*****************************************************************/
	obj.setFormat = function(format, index){
		this._formats = this._formats.concat();
		this._formats[index] = format;
	};
/****************************************************************
	Allows to specify the formatting objects for each column.
	@param formats (Array) The array of formatting objects.
*****************************************************************/
	obj.setFormats = function(formats){
		this._formats = formats;
	};
/****************************************************************
	Returns the number of the data rows.
*****************************************************************/
	obj.getCount = function(){
		if (!this._items) {return 0}
		return this._items.length;
	};
/****************************************************************
	Returns the index.
*****************************************************************/
	obj.getIndex = function(i){
		return i;
	};
/****************************************************************
	Returns the cell text.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getText = function(i, j){
		var node = this.getNode(i, j);
		var data = node ? node.text : "";
		var format = this._formats[j];
		return format ? format.dataToText(data) : data;
	};
/****************************************************************
	Returns the cell image.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getImage = function(){
		return "none";
	};
/****************************************************************
	Returns the cell hyperlink.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getLink = function(){
		return "";
	};
/****************************************************************
	Returns the cell value.
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getValue = function(i, j){
		var node = this.getNode(i, j);
		var text = node ? node.text : "";
		var format = this._formats[j];
		if (format) {
			return format.dataToValue(text);
		}
		var value = Number(text.replace(/[ ,%\$]/gi, "").replace(/\((.*)\)/, "-$1"));
		return isNaN(value) ? text.toLowerCase() + " " : value;
	};
/****************************************************************
	Returns the cell XML node text (internal).
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getNode = function(i, j){
		if (!this._items || !this._items[i]) {
			return null;
		}
		if (this._valuesPath[j]) {
			return this._items[i].selectSingleNode(this._valuesPath[j]);
		}
		else {
			return this._items[i].selectNodes(this._valuePath)[j];
		}
	};
/****************************************************************
	Returns the cell XML node text (obsolete, don't use).
	@param i (Index) Row index.
	@param j (Index) Column index.
*****************************************************************/
	obj.getData = function(i, j){
		if (!this._items) {return null}
		var node = null;
		if (this._valuesPath[j]) {
			node = this._items[i].selectSingleNode(this._valuesPath[j]);
		}
		else {
			node = this._items[i].selectNodes(this._valuePath)[j];
		}
		return node ? node.text : null;
	};
};
Active.XML.Table.create();
/******************** Grid.Paging.js  ********************/
if (!Active.Rows) {Active.Rows = {}}
Active.Rows.Page = Active.System.Model.subclass();
Active.Rows.Page.create = function(){
	var obj = this.prototype;
	obj.defineProperty("count", function(){return this.$owner.getProperty("data/count")});
	obj.defineProperty("index", function(i){return i});
	obj.defineProperty("order", function(i){return this._orders ? this._orders[i] : i});
	obj.defineProperty("text", function(i){return this.getOrder(i) + 1});
	obj.defineProperty("image", "none");
	obj.defineProperty("pageSize", 10);
	obj.defineProperty("pageNumber", 0);
	obj.defineProperty("pageCount", function(){return Math.ceil(this.getCount()/this.getPageSize())});
	var getValue = function(i){
		var size = this.getPageSize();
		var number = this.getPageNumber();
		var offset = size * number;
		return this._sorted ? this._sorted[offset + i] : offset + i;
	}
	obj.defineProperty("value", getValue);
	var getValues = function(){
		var size = this.getPageSize();
		var number = this.getPageNumber();
		var offset = size * number;
		var count = this.getCount();
		var max = count > size + offset ? size : count - offset;
		var i, values = [];
		if (this._sorted){
			values = this._sorted.slice(offset, offset + max);
		}
		else {
			for(i=0; i<max; i++){
				values[i] = i + offset;
			}
		}
		return values;
	}
	obj.defineProperty("values", getValues);
	obj.sort = function(index){
		var i, count = this.getCount();
		if (!this._sorted){
			this._sorted = [];
			for(i=0; i<count; i++){
				this._sorted[i] = i;
			}
		}
		var a = {}, direction = "ascending";
		var rows = this._sorted;
		if (this.$owner.getSortProperty("index") == index) {
			if (this.$owner.getSortProperty("direction") == "ascending") {direction = "descending"}
			rows.reverse();
		}
		else {
			for (i=0; i<rows.length; i++) {
				var text = "" + this.$owner.getDataProperty("value", rows[i], index);
				var value = Number(text.replace(/[ ,%\$]/gi, "").replace(/\((.*)\)/, "-$1"));
				a[rows[i]] = isNaN(value) ? text.toLowerCase() + " " : value;
			}
			rows.sort(function(x,y){return a[x] > a[y] ? 1 : (a[x] == a[y] ? 0 : -1)});
		}
		this._sorted = rows;
		this._orders = [];
		for(i=0; i<rows.length; i++){
			this._orders[rows[i]] = i;
		}
		this.setPageNumber(0);
		this.$owner.setSortProperty("index", index);
		this.$owner.setSortProperty("direction", direction);
	}
}
Active.Rows.Page.create();
