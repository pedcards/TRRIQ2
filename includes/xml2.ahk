﻿#Requires AutoHotkey v2.0

class XML
{
/*	new() = return new XML document
	addElement() = append new element to node object
	insertElement() = insert new element above node object
	getText() = return element text if present
	findXPath() = return xpath to node (needs work)
	saveXML() = saves XML with filename param or original filename
*/
	__New(src:="") {
		this.doc := ComObject("Msxml2.DOMDocument")
		if (src) {
			if (src ~= "s)^<.*>$") {
				this.doc.loadXML(src)
			} 
			else if FileExist(src) {
				this.doc.load(src)
				this.filename := src
			}
		} else {
			src := "<?xml version=`"1.0`" encoding=`"UTF-8`"?><root />"
			this.doc.loadXML(src)
		}
	}

	__Call(method, params) {
		if !ObjHasOwnProp(XML,method) {
			try {
				return this.doc.%method%(params[1])
			}
			catch as err {
				MsgBox("Error: " err.Message)
				return false
			} 
			}
	}

	addElement(node,child,params*) {
	/*	Appends new child to node object
		Node can be node object or XPATH
		Params:
			text gets added as text
			@attr1='abc', trims outer '' chars
			@attr2='xyz'
	*/
		node := this.isNode(node)
		try {
			IsObject(node)
		} 
		catch as err {
			MsgBox("Error: " err.Message)
			return false
		} 
		else {
			n := this.doc
			newElem := n.createElement(child)
			for p in params {
				if IsObject(p) {
					for key,val in p.OwnProps() {
						newElem.setAttribute(key,val)
					}
				} else {
					newElem.text := p
				}
			}
			node.appendChild(newElem)
			n := ""
		}
	}

	insertElement(node,new,params*) {
	/*	Inserts new sibling above node object
		Object must have valid parentNode
	*/
		node := this.isNode(node)
		try {
			IsObject(node.ParentNode)
		}
		catch as err {
			MsgBox("Error: " err.Message)
		} 
		else {
			n := this.doc
			newElem := n.createElement(new)
			for p in params {
				if IsObject(p) {
					for key,val in p.OwnProps() {
						newElem.setAttribute(key,val)
					} 
				} else {
					newElem.text := p
				}
			}
			node.parentNode.insertBefore(newElem,node)
			n := ""
		}
	}

	getText(node) {
	/*	Checks whether node exists to fetch text
		Prevents error if no text present
	*/
		node := this.isNode(node)
		try {
			return node.text
		} catch {
			return ""
		}
	}

	saveXML(fname:="") {
	/*	Saves XML
		to fname if passed, otherwise to original filename
	*/
		if (fname="") {
			fname := this.filename
		}
		this.doc.save(fname)
	}

	findXPath(node) {
	/*	Returns xpath of node
	*/
		; x := node.nodeType
		build := ""

		while (node.parentNode) {
			switch node.nodeType {
				case 1:																	; 1=Element
				{
					index := this.elementIndex(node)
					build := "/" node.nodeName "[" index "]" . build
					node := node.parentNode
				} 
				case 2:																	; 2=Attribute
				{

				}
				case 3:																	; 3=Text
				{

				}
				default:
					
			}
		}
		return build
	}

/*	====================================================================================
	INTERNAL SUPPORT FUNCTIONS
*/
	isNode(node) {
		if (node is String) {
			node := this.doc.selectSingleNode(node)
		}
		return node
	}
	elementIndex(node) {
		parent := node.parentNode
		for candidate in parent.childNodes {
			if (candidate.nodeName=node.nodeName) {
				return A_Index
			}
		}
	}
}