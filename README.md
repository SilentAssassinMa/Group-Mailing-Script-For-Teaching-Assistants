# 群发邮件脚本

### 脚本功能

这是一个用于助教给学生群发邮件的python脚本(GroupMail.py)，包括了从总表中生成个人的作业成绩并添加到邮件附件的功能(IndividualFileGenerate.py)。表格的格式参见 MailList.xlsx 和 HomeworkGrade.xlsx。

调用 IndividualFileGenerate.py 后会自动生成名为 IndividualFiles 的文件，每个人的作业成绩为以学号为名的独立文件，GroupMail.py 可以直接根据 MailList.xlsx 中的学号匹配相应文件作为邮件附件发送，如不需要这个功能，请将 doGenerateIndividualFiles 的值设为 False。

### 注意事项

1、这里默认了将 HomeworkGrade.xlsx的第一列的值作为每个文件的名字，同时默认 MailList.xlsx 的前三列分别为学号、姓名和邮箱地址，如有改动，请在代码中做相应改动。

2、在群发邮件前，请务必用自己的邮箱对脚本做一次测试，以免差错。

3、关于如何改成自己的邮箱地址，请参考代码中的注释。

4、有除此之外的任何意见和疑惑，请联系作者：silentassassin@mail.ustc.edu.cn
