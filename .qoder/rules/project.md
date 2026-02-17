---
trigger: always_on
---


## 项目背景

我正在基于vsto+vb.net开发office ai智能体插件，期望借助大模型来提升excel/word/powerpoint的能力，请你熟读根目录下的AGENTS.md和.qoder目录下过去的实现，了解我的架构和分层逻辑。

## 重要
1. 总是使用中文和我交互，我在windows系统中，所以命令行请使用powershell
2. 每次执行完成任务必须调用MS命令编译每个模块
3. 编译ShareBase模块时，需要编译整个解决方案
4. 每次新增文件都要考虑csproj这个文件的引用关系
5. 第三方库需要考虑nuget引用，所以你要能自己添加nuget引用
6. 你在解决问题的时候不要为了掩盖问题而写一堆垃圾代码
7. 每个class的行数不应该超过500行
8. 每个包的划分理一些
9. 代码行数过多时，你不用在交互面板中显示所有类代码，只用显示差异部分即可

## 项目说明
请你一定要读取AGENTS.md，他是项目的描述文件，在根目录一级ExcelAi、PowerPointAi、WordAi、ShareRibbon下都有。
该项目有4个模块，ExcelAi是负责Excel的插件，WordAi是给Word的插件，PowerPointAi是给PowerPoint使用的插件，这3个模块是实现模块，都有自己的Ribbon，他们都集成自ShareRibbon模块，ShareRibbon模块是一个抽象和公共逻辑处理模块，因为以上3个套件都有类似的Ribbon区，我就单独做了这个ShareRibbon。
ShareRibbon中有2个及其重要的类，分别是Controls/BaseChatControl.vb和Ribbon/BaseOfficeRibbon.vb，他们都有3个子类在各子模块，另外还有1个Resources/chat-template-refactored.html，这个是我右侧的用户使用的panel面板，里面存放了和vb交互的各种js代码，大模型的结果也会写到html中，通过markdown转换成html。


## 项目预期

当前已实现，大模型配置，大模型聊天，执行大模型生成的vba代码，排版，审阅功能

## 代码原则

高内聚，低耦合。
封装优先，避免单文件过大。
熟悉vsto和office机制，避免语法错误，运行时异常。
熟悉webbrowser机制。