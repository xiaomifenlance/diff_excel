# diff_excel
对比两个excel中重复的行，并输出结果

打开\diff_excel1.0\dist目录下面的 diff_excel1.0.exe 可执行文件

注意输入的路径必须为绝对路径，如果不清楚文件名的绝对路径，右键需要进行对比的文件，属性--安全--对象名称
如：

请输入源文件路径：C:\Users\Administrator\Desktop\源文件.xls

需要比对的表格编号（输入1则为第一张表，默认为第一张表）:1      # 此处是选择源文件excel中的哪一sheet表

请输入目标文件的路径:C:\Users\Administrator\Desktop\目标文件.xls


需要比对的表格编号（输入1则为第一张表，默认为第一张表）:1      # 此处是选择目标excel中的哪一sheet表

回车后开始对比...

对比结束后会在cmd 界面显示出对比结果，并在\diff_excel1.0\dist 目录下生成一个.txt文件，以时间命名

如：Result_18_12_02_09_53_43.txt
