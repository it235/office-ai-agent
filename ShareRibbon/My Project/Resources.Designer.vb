﻿'------------------------------------------------------------------------------
' <auto-generated>
'     此代码由工具生成。
'     运行时版本:4.0.30319.42000
'
'     对此文件的更改可能会导致不正确的行为，并且如果
'     重新生成代码，这些更改将会丢失。
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    '此类是由 StronglyTypedResourceBuilder
    '类通过类似于 ResGen 或 Visual Studio 的工具自动生成的。
    '若要添加或移除成员，请编辑 .ResX 文件，然后重新运行 ResGen
    '(以 /str 作为命令选项)，或重新生成 VS 项目。
    '''<summary>
    '''  一个强类型的资源类，用于查找本地化的字符串等。
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Public Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  返回此类使用的缓存的 ResourceManager 实例。
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Public ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("ShareRibbon.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  重写当前线程的 CurrentUICulture 属性，对
        '''  使用此强类型资源类的所有资源查找执行重写。
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Public Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property about() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("about", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property aiapiconfig() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("aiapiconfig", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property chat() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("chat", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找类似 &lt;!DOCTYPE html&gt;
        '''&lt;html&gt;
        '''&lt;head&gt;
        '''    &lt;meta charset=&quot;GBK&quot;&gt;
        '''    &lt;title&gt;Excel Ai Chat Content&lt;/title&gt;
        '''&lt;!-- 先加载核心库 --&gt;
        '''&lt;script src=&quot;https://officeai.local/js/highlight.min.js&quot;&gt;&lt;/script&gt;
        '''&lt;link rel=&quot;stylesheet&quot; href=&quot;https://officeai.local/css/github.min.css&quot;&gt;
        '''&lt;script src=&quot;https://officeai.local/js/marked.min.js&quot;&gt;&lt;/script&gt;
        '''&lt;script src=&quot;https://officeai.local/js/vbscript.min.js&quot;&gt;&lt;/script&gt;
        '''
        '''    &lt;script&gt;
        '''        hljs.registerAliases(&apos;vba&apos;, { languageName: &apos;vbscript&apos; });
        '''        hljs.highlightAll();
        '''    &lt;/ [字符串的其余部分被截断]&quot;; 的本地化字符串。
        '''</summary>
        Public ReadOnly Property chat_template() As String
            Get
                Return ResourceManager.GetString("chat_template", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property clear() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("clear", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property deepseek() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("deepseek", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找类似 pre code.hljs{display:block;overflow-x:auto;padding:1em}code.hljs{padding:3px 5px}/*!
        '''  Theme: GitHub
        '''  Description: Light theme as seen on github.com
        '''  Author: github.com
        '''  Maintainer: @Hirse
        '''  Updated: 2021-05-15
        '''
        '''  Outdated base version: https://github.com/primer/github-syntax-light
        '''  Current colors taken from GitHub&apos;s CSS
        '''*/.hljs{color:#24292e;background:#fff}.hljs-doctag,.hljs-keyword,.hljs-meta .hljs-keyword,.hljs-template-tag,.hljs-template-variable,.hljs-type,.hljs-variable.language_{color:#d73a49}. [字符串的其余部分被截断]&quot;; 的本地化字符串。
        '''</summary>
        Public ReadOnly Property github_min() As String
            Get
                Return ResourceManager.GetString("github_min", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查找类似 /*!
        '''  Highlight.js v11.7.0 (git: 82688fad18)
        '''  (c) 2006-2022 undefined and other contributors
        '''  License: BSD-3-Clause
        ''' */
        '''var hljs=function(){&quot;use strict&quot;;var e={exports:{}};function n(e){
        '''return e instanceof Map?e.clear=e.delete=e.set=()=&gt;{
        '''throw Error(&quot;map is read-only&quot;)}:e instanceof Set&amp;&amp;(e.add=e.clear=e.delete=()=&gt;{
        '''throw Error(&quot;set is read-only&quot;)
        '''}),Object.freeze(e),Object.getOwnPropertyNames(e).forEach((t=&gt;{var a=e[t]
        ''';&quot;object&quot;!=typeof a||Object.isFrozen(a)||n(a)})),e}
        '''e.exports=n,e.exports.default=n [字符串的其余部分被截断]&quot;; 的本地化字符串。
        '''</summary>
        Public ReadOnly Property highlight_min() As String
            Get
                Return ResourceManager.GetString("highlight_min", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property magic() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("magic", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找类似 /**
        ''' * marked v15.0.7 - a markdown parser
        ''' * Copyright (c) 2011-2025, Christopher Jeffrey. (MIT Licensed)
        ''' * https://github.com/markedjs/marked
        ''' */
        '''!function(e,t){&quot;object&quot;==typeof exports&amp;&amp;&quot;undefined&quot;!=typeof module?t(exports):&quot;function&quot;==typeof define&amp;&amp;define.amd?define([&quot;exports&quot;],t):t((e=&quot;undefined&quot;!=typeof globalThis?globalThis:e||self).marked={})}(this,(function(e){&quot;use strict&quot;;function t(){return{async:!1,breaks:!1,extensions:null,gfm:!0,hooks:null,pedantic:!1,renderer:null,silent:!1,tokenizer:null,wa [字符串的其余部分被截断]&quot;; 的本地化字符串。
        '''</summary>
        Public ReadOnly Property marked_min() As String
            Get
                Return ResourceManager.GetString("marked_min", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property mcp1() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("mcp1", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property send32() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("send32", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查找类似 /*! `vbscript` grammar compiled for Highlight.js 11.7.0 */
        '''(()=&gt;{var e=(()=&gt;{&quot;use strict&quot;;return e=&gt;{
        '''const t=e.regex,r=[&quot;lcase&quot;,&quot;month&quot;,&quot;vartype&quot;,&quot;instrrev&quot;,&quot;ubound&quot;,&quot;setlocale&quot;,&quot;getobject&quot;,&quot;rgb&quot;,&quot;getref&quot;,&quot;string&quot;,&quot;weekdayname&quot;,&quot;rnd&quot;,&quot;dateadd&quot;,&quot;monthname&quot;,&quot;now&quot;,&quot;day&quot;,&quot;minute&quot;,&quot;isarray&quot;,&quot;cbool&quot;,&quot;round&quot;,&quot;formatcurrency&quot;,&quot;conversions&quot;,&quot;csng&quot;,&quot;timevalue&quot;,&quot;second&quot;,&quot;year&quot;,&quot;space&quot;,&quot;abs&quot;,&quot;clng&quot;,&quot;timeserial&quot;,&quot;fixs&quot;,&quot;len&quot;,&quot;asc&quot;,&quot;isempty&quot;,&quot;maths&quot;,&quot;dateserial&quot;,&quot;atn&quot;,&quot;timer&quot;,&quot;isobject&quot;,&quot;filter&quot;,&quot;weekday&quot;,&quot;datevalue&quot;,&quot;c [字符串的其余部分被截断]&quot;; 的本地化字符串。
        '''</summary>
        Public ReadOnly Property vbscript_min() As String
            Get
                Return ResourceManager.GetString("vbscript_min", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查找 System.Drawing.Bitmap 类型的本地化资源。
        '''</summary>
        Public ReadOnly Property wait() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("wait", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
    End Module
End Namespace
