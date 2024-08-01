# Gppt使用简单说明
一个简单的MD文件转PPTX代码，MD文件与PPTX模板文件越规范，转化效果越好。
## 规划化示例
### MD规范化格式
最小支持四级标题
一级标题生成封面页，二级标题生成章节页，三级标题生成内容页，四级标题生成每页小标
一级二级三级标题，每一个生成一页幻灯片
目录页下，提报者与提报时间，都以四级标题形式设计
~~~
# 封面
## 目录
## 章节目录一
章节目录直属内容板块
### 内容页标题
内容页标题直属内容板块
#### 内容页下小标一
内容页下小标一内容
#### 内容页下小标二
内容页下小标二内容
---过渡页---
过渡页内容
## 章节目录二
## 章节目录三
~~~
### PPT母版布局&占位符命名规范
#### 母版布局命名：
Cover_001
Toc_001
Chapter_001
Substance_01_001 （00表示小标题数量，001表示模板序号）
Translate_001
#### 占位符命名
Cover（封面页）
	Title
	Subtitle01（副标）
	Subtitle02（提报者）
	Subtitle03（提报时间）
Toc（目录页）
	Content01
	Image01
Chapter（章节页）
	Content01
	Subtitle01
	Subtitle02
	……
Substance（内容页）
	Content01（主内容占位符）
	Image01（图片占位符）
	Subtitle01（副标占位符）
	Subcontent01（副标内容占位符）
	Subimage01（副标图片占位符）
	Subtitle02
	Subcontent02
	subimage02
	……
Translate（过渡页）
	Content01
	Image01