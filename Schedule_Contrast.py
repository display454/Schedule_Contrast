import pandas as pd
import re
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False 

def read_xls(location):
    df=pd.read_excel(location,header=2,index_col=0)
    return df

class course_info():
    def __init__(self,Name,Teacher,Time,Date,Week):
        setattr(self,Name,Name)
        self.courseName=Name
        self.courseTeacher=Teacher
        self.courseTime=Time
        self.courseDate=Date
        self.courseWeek=Week
    def courseFormat(self):
        self.courseTime=str(self.courseTime)
        self.courseTime=self.courseTime.split('-')
        self.courseTime=[int(time) for time in self.courseTime]
        # dic={"星期一":1,"星期二":2,"星期三":3,"星期四":4,"星期五":5,"星期六":6,"星期日":7}
        # self.courseDate=dic[str(self.courseDate)]
        self.courseWeek=str(self.courseWeek).split(",")
        temList=[]
        for ele in self.courseWeek:
            if len(ele)<=2:
                temList.append(int(ele))
            else:
                numa=int(ele.split("-")[0])
                numb=(int(ele.split("-")[-1]))
                temList=temList+list(range(numa,numb+1))
        self.courseWeek=temList

        
    def courseOut(self):
        print(f"课程名称：{self.courseName}")
        print(f"课程教师：{self.courseTeacher}")
        print(f"课程时间：{self.courseTime}")
        print(f"课程日期：{self.courseDate}")
        print(f"课程周次：{self.courseWeek}")
        print("*"*40)

def df_produce(df):
    courseList=[]
    for index,row in df.iterrows():
        for col in df.columns:
            content=row[col]
            if pd.isna(content) == None:
                continue
            content=str(content)
            print(content)
            matchContent=re.findall(r"(\S+)\n(\S+)\n(\S+)\(\[周\]\)\[(\S+)节\]", content)
            if matchContent != []:
                for i in range(len(matchContent)):
                    courseList.append(course_info(matchContent[i][0],matchContent[i][1],matchContent[i][3],col,matchContent[i][2]))
    for element in courseList:
        element.courseFormat()
    return courseList

class classSchedule():
    def __init__(self,list,major):
        self.schedule=list
        self.major=major
        self.cooked_schedule=[]
        self.bool_schedule=[]
        for i in range(1,21):
            weekdf=pd.DataFrame(index=range(1,13),columns=["星期一","星期二","星期三","星期四","星期五","星期六","星期日"])
            for element in list:
                if i in element.courseWeek:
                    for j in element.courseTime:
                        weekdf[element.courseDate][j]=element.courseName
            weekdf.fillna("", inplace=True)
            self.cooked_schedule.append(weekdf)
        for element in self.cooked_schedule:
            element=element.where(element=="",other=1)
            element=element.where(element!="",other=0)
            self.bool_schedule.append(element)

    def outTable(self):
        fig,axes=plt.subplots(5,4,figsize=(30,20))
        fig.suptitle(f"{self.major}专业课程表",fontsize=20)

        for i,ax in zip(range(1,21),axes.flatten()):
            ax.axis('off')
            ax.set_title(f"第{i}周")
            table=ax.table(
                cellText=self.cooked_schedule[i-1].values,
                colLabels=self.cooked_schedule[i-1].columns,
                rowLabels=self.cooked_schedule[i-1].index,
                cellLoc='center',
                loc='center',
                colColours=['#a5d8ff']*7,
                rowColours=['#a5d8ff']*12,
                cellColours=[['#f0f0f0' if cell == "" else '#d9f2e6' for cell in row] for row in self.cooked_schedule[i-1].values]
            )
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1.1,1)
        plt.show()
    def contrast(self,other):
        former=self.bool_schedule
        result_schedule=[]
        for element in other:
            for df1,df2 in zip(former,element.bool_schedule):
                result_df=df1|df2
                result_schedule.append(result_df)
            former=result_schedule
            result_schedule=[]
        
        fig,axes=plt.subplots(5,4,figsize=(30,20))
        fig.suptitle(f"空闲时间表(绿色代表共同的空闲时间)",fontsize=20)

        for i,ax in zip(range(1,21),axes.flatten()):
            ax.axis('off')
            ax.set_title(f"第{i}周")
            table=ax.table(
                colLabels=self.cooked_schedule[i-1].columns,
                rowLabels=self.cooked_schedule[i-1].index,
                cellLoc='center',
                loc='center',
                colColours=['#a5d8ff']*7,
                rowColours=['#a5d8ff']*12,
                cellColours=[['#f0f0f0' if cell == 1 else '#32CD32' for cell in row] for row in former[i-1].values]
            )
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1.1,1)
        plt.show()
        return former

a=classSchedule(df_produce(read_xls("xls文件路径")),"专业名称")
b=classSchedule(df_produce(read_xls("xls文件路径")),"专业名称")
a.outTable()
b.outTable()
a.contrast([b])
