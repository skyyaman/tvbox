#20261110 发现BUG，室内表，从中间场室打印时发现总是从第1试室开始，已解决
#20261015 发现BUG，室外表，从中间场室打印时发现总是从第1试室开始，所以Index1这里改为roomid-1，已解决
#20260629 为了节约纸张，每个试室是4科时，纵向打印，可以节约40%
#20250507 小修改，检查试室表人数是否正确
#20251226 大改:加入了特殊考生放在最后试室。优化随机排位
#20250423 升级:考务室打印时，两个试室一张A4纸打印，节省啊。去掉打印台角号的id-1的操作，避免座位号不增加问题
#20250113 升级:把班级打印的宽度设为自适应
#20241220 升级：解决考场安排时有空行的问题，解决考场安排表中间缺少考场编号问题，考场编号可以不连续了
#20241118 发现BUG，291行，试室号和试室地址出现重复要先检查。已修复
#20240429修改，修改考务室打印，把各科组合一起，放入一张纸进行打印，太浪费了 2024/4/29
#20231221修改，发现上次1024的修改没能完全修复，中间人数为0还是不能正常运行，现在可以了，然后还加了一个室外表多人的时候可以选择加在哪一列
#20231128修正，班级表打印不下，调整了列宽以便于一页打印。
#20231115发现bug未修，当单独打印最后试外表时，不能输出正确的人数
#20231024修改，修复中间试室人数为0时不能运行之bug
#20230616修改，修复选科单独考试时排序还是以组合选科来排序的问题。现在选科单独考时排序正常了
#20230206修改。考号列以文本形式读取，以前以数字读取的，以0开头就SB了
#20230205修改。加上室内表空白课桌的输出
#20230105大改。修复了中间空白班级时打印出错的问题，很多问题.lotlotlotof questions
#2022/11/14修改。加上了物历混合考场 2022/12/5修改，改正了班级打印表多次打印出错的BUG
#打包：pyinstaller --add-data "open.gif;." -F -w -i logo.ico --console examroom.py
#佛山的考试选科是高考模式，要统计各选科的地址，用此公式可以：=INDEX($M$2:$X$758,ROW(1:1),MATCH("*" & Y2 & "*",$M$1:$X$1,0)+1)
import tkinter as tk
import xlsxwriter
import pandas as pd
import requests
import os,sys
import re
import time
import threading
import xlwings as xw
from win32.win32api import GetSystemMetrics
from win32com.client import Dispatch, constants, gencache, DispatchEx
from win32 import  win32gui, win32print
from win32.lib import win32con
from tkinter import  messagebox,filedialog,IntVar
import ttkbootstrap as ttk
from ttkbootstrap.constants import *    #主题
try:
    pd.set_option('future.no_silent_downcasting', True)
except Exception:
    pass
__version__ = "26.11.10"

class anpai(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.special_ids = []   # 用于存储特殊考号                
        self.title(u'试室安排')
        self.geometry("%dx%d+%d+%d" % (x,y,(self.winfo_screenwidth()-x)/2+x+2,(self.winfo_screenheight()-y)/2))
        self.attributes("-toolwindow", 2)
        self.resizable(False,False)        
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.setup_UI()

    def setup_UI(self):
        global flagWin1
        flagWin1=True
        # 对话框界面        
        canvas = tk.Canvas(self,width=x,height=y) #创建一个画布
        canvas.pack()
        canvas.create_line(0,100,x,100,fill="dark green")         #画个直线
        tk.Label(self, text='文件位置：',bg='orange',fg='white',font=('microsoft yahei',10)).place(x=10,y=30, width=70, height=25) 
        self.openfile1= ttk.Entry(self,width=50)
        self.openfile1.insert(0,fileurl)
        self.openfile1.place(x=90,y=30, width=250, height=25)

        b=ttk.Button(self,bootstyle='info')
        b.place(x=350,y=30, width=30, height=25)
        b.focus()

        tk.Label(self, text='导出安排文件名：',bg='orange',fg='white',font=('microsoft yahei',10)).place(x=10,y=60, width=120, height=25) 
        self.targetLocation=ttk.Entry(self,width=200)  #默认输出文件名
        self.targetLocation.insert(0,'temp_'+examname)
        self.targetLocation.place(x=140,y=60, width=200, height=25)
        tk.Label(self, text='.xlsx',fg='red').place(x=341,y=60, width=30, height=25)
        
        self.loadfile = ttk.Button(self, text='加载考场安排文件',command=self.LoadFile,bootstyle='primary').place(x=10,y=120, width=120, height=30)#加载excel
        self.closeAnpai = ttk.Button(self, text='关闭此窗口',command=self.exit1,bootstyle='info').place(x=160,y=300, width=100, height=30)
        b.bind('<ButtonRelease-1>',self.open)

    def delStrFromList(self,l):
        n=len(l)
        bek=[]
        for i in range(n):
            if "Unnamed" in str(l[i]):
                bek.append(i)                
        a = [l[i] for i in range(n) if (i not in bek)]
        return a
    def open_special_window(self):
        win = tk.Toplevel(self)
        win.title("特殊考生考号")
        win.resizable(False, False)

        win.update_idletasks()  # 先让窗口计算尺寸

        w, h = 400, 300
        x = self.winfo_x() + (self.winfo_width() - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2

        win.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(win,text="请粘贴特殊考生考号（每行一个）：",font=('microsoft yahei', 10)).pack(pady=5)
        text = tk.Text(win, width=45, height=12)
        text.pack(padx=10)

        def save_special():
            content = text.get("1.0", "end").strip()
            if not content:
                self.special_ids = []
            else:
                # 每行一个，去空格、去空行、去重
                self.special_ids = list({
                    line.strip()
                    for line in content.splitlines()
                    if line.strip()
                })
            # ===== 校验特殊考生考号是否存在 =====
            if hasattr(self, 'special_ids') and self.special_ids:
                all_ids = set(self.StData['考号'].astype(str))
                invalid_ids = [i for i in self.special_ids if i not in all_ids]

                if invalid_ids:
                    messagebox.showerror(
                        "特殊考生考号错误",
                        "以下特殊考生考号不在学生数据中：\n\n"
                        + "\n".join(invalid_ids)
                    )
                    return
            messagebox.showinfo("完成",f"已读取 {len(self.special_ids)} 个特殊考生考号")
            win.destroy()

        ttk.Button(win,text="保存", command=save_special,bootstyle="success").pack(pady=10)
    def randclass(self, data, Subjectlist): 

        # --- 1. 预处理科目列表，提取不重复的科目/组合字符串 ---
        newlist1 = []
        for i in Subjectlist:
            parts = i.split('+') if '+' in i else [i]
            for p in parts:
                # 提取非数字字符，例如 '生地10' -> '生地'
                newstring = ''.join([j for j in p if not j.isdigit()])
                if newstring and newstring not in newlist1:
                    newlist1.append(newstring)

        # --- 2. 核心逻辑：建立识别列，解决类型冲突 ---
        # 使用 '科目分组' 处理数字累加，'科目展示' 处理字符串名称
        data['科目分组'] = 0
        data['科目展示'] = ""

        for i in newlist1:
            if len(i) == 1:
                # 单场科目直接赋值
                mask = data['选科'].str.contains(i, na=False)
                data.loc[mask, '科目展示'] = i
                # 给单场分配一个足够大的特征值，避免与混合场的数字碰撞
                # 使用 ord(i) 获取字符编码作为唯一ID
                data.loc[mask, '科目分组'] = ord(i) 
            else:
                # 针对混合场：生、化、政、地
                if any(k in i for k in '生化政地'):
                    # 累加权重：生1, 化2, 政4, 地8
                    data.loc[data['选科'].str.contains('生', na=False), '科目分组'] += 1
                    data.loc[data['选科'].str.contains('化', na=False), '科目分组'] += 2
                    data.loc[data['选科'].str.contains('政', na=False), '科目分组'] += 4
                    data.loc[data['选科'].str.contains('地', na=False), '科目分组'] += 8
                    # 这里的 '科目展示' 会在最后根据分组统一生成，现在先存组合名
                    mask = data['选科'].str.contains(i, na=False)
                    data.loc[mask, '科目展示'] = i
                else:
                    # 其他组合（如物历、语数）
                    mask = data['选科'].str.contains(i, na=False)
                    data.loc[mask, '科目展示'] = i
                    # 用字符串哈希值做分组ID
                    data.loc[mask, '科目分组'] = hash(i)

        # 统一将处理好的识别信息存回 '科目' 列，此时全部转为字符串，确保后续排序稳定
        data['科目'] = data['科目分组'].astype(str)
        
        # 剔除未分配到考场科目的学生（如果有）
        data = data[data['科目分组'] != 0].copy()
        
        # --- 3. 班级打散算法优化 ---
        # 我们按科目分组进行循环，在每组内尽可能让相邻的人班级不同
        sub_groups = data['科目'].unique()
        final_list = []

        for j in sub_groups:
            datasplit = data[data['科目'] == j].copy()
            if datasplit.empty:
                continue
                
            # 初始取第一行
            current_data = datasplit.iloc[[0]]
            datasplit = datasplit.drop(current_data.index)
            group_res = [current_data]
            
            last_class = current_data.iloc[0]['班级']
            
            # 循环抽取后续学生
            while not datasplit.empty:
                # 寻找与上一人班级不同的学生
                diff_class_mask = datasplit['班级'] != last_class
                rest = datasplit[diff_class_mask]
                
                if not rest.empty:
                    # 随机抽取一个不同班级的学生
                    next_st = rest.sample(n=1)
                    last_class = next_st.iloc[0]['班级']
                else:
                    # 如果剩下的全是同一个班的，就直接取剩下的第一个
                    next_st = datasplit.iloc[[0]]
                    last_class = next_st.iloc[0]['班级']
                
                group_res.append(next_st)
                datasplit = datasplit.drop(next_st.index)
                
            final_list.extend(group_res)

        # 使用 pd.concat 一次性合并，效率比循环 concat 高几百倍
        return pd.concat(final_list, axis=0).reset_index(drop=True) 
    def randclass111(self,data,Subjectlist): ##去掉每科目的重复['英','日'] [物10+历5，历]
        #下面看怎么对选科列进行排序，按subjectlist
        newlist=[]
        templist=[]
        
        data['科目']=0  #原来用这一行会引起警告，用下面消除警告
        
        for i in Subjectlist:
            if '+' in i:    #混合考场
                templist=i.split('+')
                for k in templist:
                    newstring = ''.join([j for j in k if not j.isdigit()])  #得到 '生地10'中的生地字符，去掉数字
                    newlist.append(newstring)   #把各科分割开放入列表
            else:
                newlist.append(i)
        #print(newlist)           newlist=['语数'] newlist=['生政', '化地', '政地', '生地', '化政', '化生']
        newlist1=[]
        [newlist1.append(i) for i in newlist if not i in newlist1]  #对newlist进行数据合并，不重复科目

        for i in newlist1:
            #当选科单独考场时，发生班级混在一起的现象，原因在于下面的科目，只有生化这样多选科在一起才用下面的方法，否则直接写科目
            if len(i)==1:
                temp1=data['选科'].str.contains(i)
                data.loc[data[temp1].index,'科目']=i
            else:
                if '生' in i or '化' in i or '政' in i or '地' in i:
                    temp1=data['选科'].str.contains('生')
                    temp2=data['选科'].str.contains('化')
                    temp3=data['选科'].str.contains('政')
                    temp4=data['选科'].str.contains('地')

                    data.loc[data[temp1].index,'科目']=data.loc[data[temp1].index,'科目']+1
                    data.loc[data[temp2].index,'科目']=data.loc[data[temp2].index,'科目']+2
                    data.loc[data[temp3].index,'科目']=data.loc[data[temp3].index,'科目']+4
                    data.loc[data[temp4].index,'科目']=data.loc[data[temp4].index,'科目']+8
                else:                    
                    temp1=data['选科'].str.contains(i)
                    data.loc[data[temp1].index,'科目']=i

        data = data.drop(data[data['科目']==0].index)
        data=data.sort_values(by=['科目'],na_position='first',ascending=(True)) #
        
        subcount=list(data['科目'].unique())    #科目列表
        newpd=pd.DataFrame(columns=data.columns)    #建立新PD，结构为data已有的
        
        for j in subcount:
            #print(j,11111)
            datasplit=data[data['科目'] == j]   #把本科的筛选出来 
            newpd=pd.concat([newpd,datasplit.iloc[[0]]])   #先得到第一行数据    
            datasplit=datasplit.drop(index=datasplit.iloc[[0]].index)
            
            for i in range(datasplit.shape[0]):
                classID=newpd.iat[newpd.shape[0]-1,1]          #班级号 这里原来用的第2列 一直是错的
                rest=datasplit[datasplit['班级'] != classID]  #班级号不相同的记录   

                if not rest.empty:  #后面有不同的班级，就取样，否则就全部合在一起了
                    nextst=rest.sample(n=1)                 #这个是从后面进行随机取样，会出重复的问题。要用何老板的算法才行。等何老板的函数
                    newpd=pd.concat([newpd,nextst.iloc[0:]],axis=0)
                    datasplit=datasplit.drop(index=nextst.index)    #从原数据里把已抽取的数据删除
                    
                else:   #后面都是一个班的
                    newpd=pd.concat([newpd,datasplit],axis=0)# 一个班的，全合在一起了
                    break   
    
        return newpd
    
    def swap_special_students(self, df, subject_col):
        import re
        if not hasattr(self, 'special_ids') or not self.special_ids:
            return df

        # --- 1. 辅助解析函数 ---
        def parse_room_desc(desc, actual_total_count):
            if pd.isna(desc): return []
            parts = str(desc).split('+')
            parsed = []
            current_offset = 0
            if '+' not in str(desc):
                chars = set(re.findall(r'[\u4e00-\u9fa5]', str(desc)))
                if chars:
                    parsed.append({'sub_set': chars, 'count': int(actual_total_count), 'end_pos': int(actual_total_count)})
            else:
                for p in parts:
                    count_match = re.search(r'(\d+)', p)
                    count = int(count_match.group(1)) if count_match else 0 
                    chars = set(re.findall(r'[\u4e00-\u9fa5]', p))
                    if chars:
                        current_offset += count
                        parsed.append({'sub_set': chars, 'count': count, 'end_pos': current_offset})
            return parsed

        # --- 2. 收集所有特殊生及其归属目标 ---
        # 记录哪些特殊生应该去哪个 (考场, 组起始座, 组结束座)
        assignments = {} 
        count_col = subject_col + '人数'

        special_indices = df[df['考号'].isin(self.special_ids)].index.tolist()
        for s_idx in special_indices:
            st_selection = set(re.findall(r'[\u4e00-\u9fa5]', str(df.at[s_idx, '选科'])))
            potential_targets = []
            for _, row_room in self.RoomData.iterrows():
                room_id = row_room['试室号']
                real_count = row_room[count_col] if count_col in row_room else 42
                room_structs = parse_room_desc(row_room[subject_col], real_count)
                for group in room_structs:
                    if group['sub_set'].issubset(st_selection):
                        potential_targets.append({'room_id': room_id, 'end': group['end_pos'], 'cnt': group['count']})
            
            if potential_targets:
                best = max(potential_targets, key=lambda x: x['room_id'])
                key = (best['room_id'], best['end'] - best['cnt'] + 1, best['end'])
                if key not in assignments: assignments[key] = []
                assignments[key].append(s_idx)

        # --- 3. 核心：强制填满末尾座位 ---
        for (room_id, start_seat, end_seat), s_indices in assignments.items():
            
            num_to_fill = len(s_indices)
            # 确定的目标座位号列表 (例如: [39, 40, 41])
            target_seat_numbers = list(range(end_seat - num_to_fill + 1, end_seat + 1))
            
            # A. 找出目前坐在这些位置上的【所有】行索引
            current_occupants_indices = df[(df[subject_col+'考场号'] == room_id) & 
                                          (df[subject_col+'座位号'].isin(target_seat_numbers))].index.tolist()
            
            # B. 找出这些位置中，哪些目前是【正常生】(需要被踢走的)
            normal_student_indices_to_be_replaced = [i for i in current_occupants_indices if df.at[i, '考号'] not in self.special_ids]
            
            # C. 找出【不在】这些目标位置的特殊生 (需要搬进来的)
            # 判断标准：考场号不对，或者座位号不在 target_seat_numbers 里
            specials_to_move_in = []
            for s_idx in s_indices:
                is_in_correct_room = df.at[s_idx, subject_col+'考场号'] == room_id
                is_in_correct_seat = df.at[s_idx, subject_col+'座位号'] in target_seat_numbers
                if not (is_in_correct_room and is_in_correct_seat):
                    specials_to_move_in.append(s_idx)

            # D. 一对一强制交换
            # 此时 specials_to_move_in 的数量一定等于 normal_student_indices_to_be_replaced
            for s_idx, n_idx in zip(specials_to_move_in, normal_student_indices_to_be_replaced):        
                
                cols = [subject_col + '地址', subject_col + '考场号', subject_col + '座位号']
                for col in cols:
                    df.at[s_idx, col], df.at[n_idx, col] = df.at[n_idx, col], df.at[s_idx, col]
        
        return df
    
    def arrange(self):
        starttime=time.time()
        #准备把上面的整合到一个函数里面来
        replace_dict = {'英语': '英', '日语': '日', '生物': '生', '物理': '物', '历史': '历', '政治': '政', '地理': '地', '化学': '化'}
        stdata=self.StData
        roomdata=self.RoomData
        STCountListNum=0
          
        self.resultDf=stdata.sort_values(by=['班级','学号'],na_position='first',ascending=(True,True))

        #检查考号的唯一性
        if stdata.shape[0]!=stdata['考号'].nunique():
            messagebox.showerror('出错','考号有重复，请检查')
            return
        colList=self.delStrFromList(list(roomdata)) #列名标签

        if len(colList) % 4!=2:
            messagebox.showinfo('出错','场室表有误，是否多写了几列？')
            return        
        colName=getname(colList)    #科目列表
        #这里要用列号
        RoomList=roomdata.iloc[:,1].values.tolist() #场室地址
        roomidlist=roomdata.iloc[:,0].values.tolist() #场室号
        for Subject in colName:  #从场室表里得到要安排的科目，循环每个科目，打乱，然后安排各自，最后合并。 
            
            STCountList=roomdata[Subject+'人数'].values.tolist() #该科人数  [nan, nan, nan, nan, nan, nan, nan, nan, 42.0, nan, nan, 42.0, 42.0, nan, 42.0, nan, 12.0, nan]
 
            SubjectList=roomdata[Subject][roomdata[Subject].notnull()].values.tolist()#选科这里有可能是空，要先过滤
            #SubjectList=roomdata[Subject].values.tolist()#选科这里有可能是空，要先过滤
            '''SubjectList=[i.replace('语文','语').replace('数学','数').replace('英语','英') \
                .replace('日语','日').replace('物理','物').replace('化学','化') \
                .replace('生物','生').replace('政治','政').replace('历史','历') \
                .replace('地理','地')  for i in SubjectList]            #对科目信息进行简单化'''
            for i in SubjectList:
                if i in replace_dict:SubjectList[SubjectList.index(i)]=replace_dict[i]
            if len(SubjectList)==0:continue  #科目列表为空，这科不用安排。PASS  ['政', '政', '政', '政', '政']
            
            df=stdata.sample(frac=1).reset_index(drop=True)     #打乱  

            self.df1=df.drop(index=df.index)    #原表清空，做个原结构的新表
            self.df1[Subject+'地址']=0             #后面加两列
            self.df1[Subject+'座位号']=0
            realsub=[]      #去掉每科目的重复['英','日'] [物10+历5，历]
            [realsub.append(i) for i in SubjectList if not i in realsub]    #NMD列表推导式。。。

            df=self.randclass(df,realsub)   #得到一个本科的不重复班级的DF
            
            STCountListNum=0
            STCountList = [0 if pd.isna(x) else x for x in STCountList]

            if STCountList[-1]==0:STCountList.pop()# 去除最后一个元素，如果它是NaN,STCountList是该科人数，列表，
 
            #for i in range(len(STCountList)):  #每一个考场
            for index,i in enumerate(STCountList):  #每一个考场
                nokey=1
                if not i :continue
                try:                    
                    SubjectChild=SubjectList[STCountListNum].split('+') #得到一个子列表，可能是1、2、3个科目，有合场的用+号连接
                    if len(SubjectChild)==1:        #本场室只有一个科目，最简单
                        
                        df = df.replace(replace_dict, regex=True) #用字典替换，下面的写法警告
                        #df=df.replace('英语','英',regex=True).replace('日语','日',regex=True).replace('生物','生',regex=True) \
                        #    .replace('物理','物',regex=True).replace('化学','化',regex=True).replace('地理','地',regex=True)
                        dfdata=df
                        for j in range(len(SubjectList[STCountListNum])):#科目中每个字都找一遍
                            is_exist_sub=df['选科'].str.contains(SubjectList[STCountListNum][j])    #选科列中符合科目的个数
                            if is_exist_sub.values.sum() or (not i):    
                                nokey=0
                                dfdata=dfdata.loc[is_exist_sub]
                        if nokey:
                            messagebox.showerror('错误','没有找到科目信息，检查两个表的科目是否对应？')
                            return
                        
                        dfdata=dfdata.iloc[0:int(i)]
                        dfdata[Subject+'地址']=RoomList[index]      #场室地址写入+
                        dfdata[Subject+'考场号']=roomidlist[index]      #考场号，此处错误，因试室号不一定连续
                        
                        try:          
                            dfdata[Subject+'座位号']=range(1,int(i)+1) #座位号
                        except:
                            messagebox.showerror('出错',f'也许是{Subject}科的学生人数不对？检查一下大侠~')
                            return
                        dfIndex=dfdata.index  #取前42个数据的index，以便在原数据中删除
                        self.df1=pd.concat([self.df1,dfdata])   #合并两个表
                        df.drop(labels=dfIndex,inplace=True)      #删除。
                        
                    else:      #有多个科目。
                        start=1   # 混合场室的开始顺序号
                        end=0                    
                        for xk in SubjectChild: #SubjectChild=['政生10', '生地22']    [历23，物10]      
                                        
                            start+=end                        
                            end=int(re.findall(r'\d+',xk)[0])  #得到人数，
                            kemu_str=re.findall(r'\D+',xk)[0]     #科目的文字长度，是1位=历/物 2位=选科
                            
                            #这个可以打乱吗
                            if len(kemu_str)==1:
                                dfdata=df.loc[df['选科'].str.contains(kemu_str)].iloc[0:end]
                            else:
                                dfdata=df.loc[df['选科'].str.contains(xk[0]) & df['选科'].str.contains(xk[1])].iloc[0:end]
                            dfdata[Subject+'地址']=RoomList[index]
                            #dfdata[Subject+'考场号']=index+1      #考场号
                            dfdata[Subject+'考场号']=roomidlist[index]      #考场号，此处错误，因试室号不一定连续
                            dfdata[Subject+'座位号']=range(start,start+end)  
                            dfIndex=dfdata.index  #取前42个数据的index     
                            self.df1=pd.concat([self.df1,dfdata])   #合并两个表
                            
                            df.drop(labels=dfIndex,inplace=True)  #删除。
                    STCountListNum+=1
                except:
                    continue
            self.resultDf=pd.merge(self.resultDf,self.df1[['考号',Subject+'地址',Subject+'考场号',Subject+'座位号']],how='outer',on=['考号'])
            self.resultDf = self.swap_special_students(self.resultDf, Subject)
        #整理resultDf的数据
        cols_to_fix = [
            '语数外考场号', '语数外座位号', 
            '物历考场号', '物历座位号', '选科考场号', '选科座位号',
            '化考场号', '化座位号', 
            '生考场号', '生座位号', 
            '政考场号', '政座位号', 
            '地考场号', '地座位号'
        ]

        # 使用 Int64 (注意是大写的 I)
        for col in cols_to_fix:
            if col in self.resultDf.columns:
                self.resultDf[col] = self.resultDf[col].astype('Int64')
        #收集特殊生信息
        logs=[]
        for i in self.special_ids:
            stline=self.resultDf[self.resultDf['考号']==i]
            subjectdic={}
            for subject in colName:
                #print(type(stline.iloc[0][subject+'考场号']),type(stline.iloc[0][subject+'座位号']))
                if pd.isnull(stline.iloc[0][subject+'考场号']):
                    continue
                subjectdic[subject] = str(stline.iloc[0][subject+'考场号'])+ '-' +str(stline.iloc[0][subject+'座位号'])
            
            base_info = {
                '学籍号': stline.iloc[0]['学籍号'],
                '考号': stline.iloc[0]['考号'],
                '班级': stline.iloc[0]['班级'],
                '学号': stline.iloc[0]['学号'],
                '姓名': stline.iloc[0]['姓名'],
            }
            combined_entry = {**base_info, **subjectdic}    
            # 将合并后的字典加入 logs
            logs.append(combined_entry)
        #except:
        #    messagebox.showerror('出错','不知道哪里出错，也许是场室的表头有错？检查一下大侠~')
        #    return
        filename=workdir+self.targetLocation.get()

        try:
            self.resultDf.to_excel(filename+'.xlsx',index=False)
            with pd.ExcelWriter(filename+'.xlsx') as writer:
                self.resultDf.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
                roomdata.to_excel(writer, sheet_name='Sheet2', index=False, header=True)
            
                if logs:
                    log_df = pd.DataFrame(logs)
                    log_df.to_excel(
                        writer,
                        sheet_name='特殊考生',
                        index=False
                    )
        except:
            self.progresstext.config(text="🤢任务出错！🤢")  # 更新文字
            self.test4.config(text="👉考场安排")
            self.progress_bar.stop()  # 停止进度条动画
            messagebox.showerror('出错','是'+filename+'文件打开未关闭吗？先关闭再说~')
            return
        wastetime=time.time()-starttime
        self.progress_bar.place_forget()  # 停止进度条动画
        self.test4.config(text="👉考场安排")
        self.progresstext.config(text="😀任务完成！😀")  # 更新文字
        messagebox.showinfo('完成','😀座位安排完毕，已写入文件'+filename+'.xlsx'+f'，程序耗时{round(wastetime,3)}秒')
    

    def start_task(self):
        self.test4.config(text="安排中...")
        progress_bar_x = 20
        progress_bar_y = 250
        progress_bar_width = 350
        self.progress_bar = ttk.Progressbar(self, mode='indeterminate')  # 不确定模式
        self.progress_bar.pack(pady=20)
        self.progress_bar.place(x=progress_bar_x,y=progress_bar_y, width=progress_bar_width, height=30)
        self.progress_bar.start(10)  # 以10ms的间隔更新进度条动画
        self.progresstext = ttk.Label(self,text="🐎快马加鞭安排中🐎",foreground="red")
        self.progresstext.place(x=progress_bar_x + progress_bar_width / 2 - 50, y=progress_bar_y +5)  # 偏移 -30 到进度条上方
        task_thread = threading.Thread(target=self.arrange)
        task_thread.start()
    def LoadFile(self):
        xl=pd.ExcelFile(self.openfile1.get())
        f=xl.sheet_names
        ttk.Label(self, text=f'excel读取完成，共{len(f)}个工作表,表名如下,确保学生和试室在前两个：').place(x=20,y=150, width=380, height=40)
        j=k=0
        l=5 #每行放5个标签
        self.btn_special = ttk.Button(self,text='👉特殊考生',command=self.open_special_window,bootstyle="warning")
        self.btn_special.place(x=140, y=120, width=100, height=30)
        self.test4 = ttk.Button(self, text='👉考场安排',command=self.start_task,bootstyle="success")
        self.test4.place(x=250,y=120, width=100, height=30)
        for i in xl.sheet_names:  
            if j==5:j=0
            if (20+68*j)>360:
                ttk.Button(self,text=i[:4],bootstyle="light").place(x=20+68*j,y=200+k//l*30, width=68, height=30)
            else:
                ttk.Button(self,text=i[:4],bootstyle="light").place(x=20+68*j,y=200+k//l*30, width=68, height=30)
            j+=1
            k+=1
        try:
            self.StData = pd.read_excel(xl, sheet_name=f[0], dtype={'考号': str})
            self.RoomData=xl.parse(sheet_name=f[1]).dropna(how="all").reset_index(drop=True)     #场室信息           
        except:
            messagebox.showerror('错误','检查文件是否正确')
            return
        if not self.RoomData.iloc[:, 0].is_unique or not self.RoomData.iloc[:, 1].is_unique: #检查前两列的唯一性
            messagebox.showerror('错误','第一二列有重复数据，请检查')
            return


    def exit1(self):
        global flagWin1
        flagWin1=False
        self.destroy()
    def open(self,event=None):
        global workdir
        self.filelocation =filedialog.askopenfilename(filetypes=[('Excel','.xls .xlsx .xlsm')])
        self.openfile1.delete(0,"end")
        self.openfile1.insert(0,self.filelocation)
        #取得工作目录
        try:
            workdir=re.findall(r'(^.+\/).',self.filelocation)[0]
        except:
            workdir=''
    def cancel(self):
        self.Retcode = None
        self.exit1()

class Win2Print(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.title(u'打印')
        self.geometry("%dx%d+%d+%d" % (x,y,(self.winfo_screenwidth()-x)/2-x-2,(self.winfo_screenheight()-y)/2))
        self.attributes("-toolwindow", 2)
        self.resizable(False,False)
        #self.wm_attributes('-topmost',True)    # 不要置顶了
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.setup_UI()

    def setup_UI(self):
        global flagWin2
        flagWin2=True
        canvas = tk.Canvas(self,width=x,height=y) #创建一个画布
        canvas.pack()
        canvas.create_line(0,100,x,100,fill="dark green")         #画个直线
        # 对话框界面
        self.myLabel = tk.Label(self, text='考试名称：',bg='orange',fg='white',font=('microsoft yahei',10)).place(x=10,y=70, width=70, height=25) 
        self.myEntryBox = ttk.Entry(self,width=50)      #这是考试名称文本框
        self.myEntryBox.focus_set()         #直接获得焦点，然后文本有改变事件则改变焦点        
        self.myEntryBox.insert(0,examname)
        self.myEntryBox.place(x=90,y=70, width=250, height=25)

        a=tk.Label(self, text='导入数据：',bg='orange',fg='white',font=('microsoft yahei',10))
        a.place(x=10,y=30, width=70, height=25) 
        self.openfile1= ttk.Entry(self,width=50) #openfile1是文件目录拾取
        self.openfile1.insert(0,printFileUrl)
        self.openfile1.place(x=90,y=30, width=250, height=25)
        #打开文件图标并绑定
        #image_file=tk.PhotoImage(file=os.path.join(__DIR__, "open.gif"),width=20,height=15)
        #b=ttk.Button(self,image=image_file,bootstyle='info')
        b=ttk.Button(self,bootstyle='info')
        b.place(x=350,y=30, width=30, height=25)
        b.focus()
        self.CheckVar1 = IntVar()
        self.text=tk.StringVar()
        _output=[[IntVar(),'导出为PDF',],[IntVar(),'导出后打印',]]
        self.output=_output
        loadfile = ttk.Button(self, text='加载安排文件',command=self.LoadFile,bootstyle="primary").place(x=10,y=110, width=120, height=30)#加载excel
        self.pdfoutput=ttk.Checkbutton(self, text=_output[0][1],variable =_output[0][0],onvalue=1,offvalue=0,bootstyle="success-round-toggle").place(x=140,y=110, width=100, height=30)
        self.synprint=ttk.Checkbutton(self, text=_output[1][1],variable =_output[1][0],onvalue =1, offvalue =0,bootstyle="success-round-toggle").place(x=250,y=110, width=120, height=30)
        
        outmsg=ttk.Label(self,bootstyle="danger",textvariable=self.text).place(x=5,y=300, width=400, height=30)
        ttk.Button(self, text='关闭此窗口',command=self.exit2,bootstyle="success").place(x=160,y=330, width=100, height=30)
        b.bind('<ButtonRelease-1>',self.open)
        
    def LoadFile(self):
        xl=pd.ExcelFile(self.openfile1.get())
        f=xl.sheet_names
        ttk.Label(self, text=f'excel读取完成，共{len(f)}个工作表,表名如下：').place(x=20,y=245, width=380, height=30)
        j=k=0
        l=5 #每行放5个标签
        #self.StData=xl.parse(sheet_name=f[0])       #学生信息
        self.StData = pd.read_excel(xl, sheet_name=f[0], dtype={'考号': str})#考号列应该是以文本形式读取，上面行以数字读，会出问题
        self.RoomData=xl.parse(sheet_name=f[1]).dropna(how="all").reset_index(drop=True)    #场室信息
        SubjectList=self.RoomData.iloc[:,1][self.RoomData.iloc[:,1].notnull()].values.tolist()
        ttk.Label(self, text='从').place(x=10,y=145, width=15, height=25)
        self.startroom = ttk.Entry(self,width=10)
        self.startroom.focus_set()         #直接获得焦点，然后文本有改变事件则改变焦点        
        self.startroom.insert(0,1)
        self.startroom.place(x=25,y=145, width=30, height=25)
        ttk.Label(self, text='到').place(x=55,y=145, width=20, height=25)
        self.endroom = ttk.Entry(self,width=30,bootstyle="primary")
        self.endroom.focus_set()         #直接获得焦点，然后文本有改变事件则改变焦点
        self.endroom.insert(0,self.RoomData.iat[-1, 0])        
        self.endroom.place(x=75,y=145, width=30, height=25)
        ttk.Label(self, text='室。').place(x=105,y=145, width=15, height=25)
        ttk.Button(self, text='室外张贴表',command=self.outSideRoom,bootstyle="success").place(x=130,y=145, width=90, height=30)
        ttk.Button(self, text='课桌张贴表',command=self.inSideRoom,bootstyle="success").place(x=235,y=145, width=90, height=30)
        ttk.Button(self, text='打印考务室',command=self.kaowu,bootstyle="success").place(x=130,y=180, width=90, height=30)
        ttk.Button(self, text='打印试卷袋',command=self.printbag,bootstyle="success").place(x=235,y=180, width=90, height=30)

        ttk.Label(self, text='从').place(x=10,y=215, width=15, height=25)
        self.startclass = ttk.Entry(self,width=10)
        self.startclass.focus_set()         #直接获得焦点，然后文本有改变事件则改变焦点        
        self.startclass.insert(0,1)
        self.startclass.place(x=25,y=215, width=30, height=25)   
        
        ttk.Label(self, text='到').place(x=55,y=215, width=15, height=25)
        self.endclass = ttk.Entry(self,width=30,bootstyle="primary")
        self.endclass.focus_set()         #直接获得焦点，然后文本有改变事件则改变焦点        
        self.endclass.insert(0,16)
        self.endclass.place(x=75,y=215, width=30, height=25)
        ttk.Label(self, text='班').place(x=105,y=215, width=20, height=25)
        ttk.Button(self, text='打印班级表',command=self.ClassTable,bootstyle="info").place(x=130,y=215, width=90, height=30)
        
        self.colList=anpai.delStrFromList(self,list(self.RoomData))
        if len(self.colList) % 4!=2:
            messagebox.showinfo('出错','场室表有误，是否多写了几列？')
            return
        colName=[]
        self.subjectSum=0
        self.RealSubject=[]
        for i in range(int(len(self.colList)/4)):colName.append(self.colList[4*i+2])  #得到科目列表        
        for Subject in colName:  #从场室表里得到要安排的科目，循环每个科目，打乱，然后安排各自，最后合并。
            SubjectList=self.RoomData[Subject][self.RoomData[Subject].notnull()].values.tolist()#选科这里有可能是空，要先过滤
            if len(SubjectList)>0:
                self.RealSubject.append(Subject)   #得到实际参考科目数

        for i in xl.sheet_names:  #考安排.xlsx]阅卷教师>
            if j==5:j=0
            if (20+68*j)>360:
                ttk.Button(self,text=i[:6],bootstyle="warning").place(x=20+68*j,y=270+k//l*30, width=68, height=30)
            else:
                ttk.Button(self,text=i[:6],bootstyle="warning").place(x=20+68*j,y=270+k//l*30, width=68, height=30)
            j+=1
            k+=1
    
    def clickselect(self,event=None):
        self.myEntryBox.delete(1,5)

    def printbag(self):
        try:
            workbook = xlsxwriter.Workbook(workdir+bag)
            stdata=self.StData
            printroomlist=list(range(int(self.startroom.get()),int(self.endroom.get())+1))   #要打印的考场号列表  
            #testroomlist=self.RoomData.iloc[:,1].values.tolist()  #所有实际考场地点列表      
            idroomlist=self.RoomData.iloc[:,0].values.tolist()    #所有实际考场号列表 
            cell_format = workbook.add_format({
                'bold':     False,
                'border':   1,
                'align':    'center',#水平居中
                'valign':   'vcenter',#垂直居中
                'font_name':u'微软雅黑',
                'font_size':10,
                'text_wrap':1
            })
            title_format = workbook.add_format({
                'bold':     True,
                'border':   0,
                'align':    'center',#水平居中
                'valign':   'vcenter',#垂直居中
                'font_name':u'微软雅黑',
                'font_size':15,
                'text_wrap':1
            })
            for RealSubjectid in range(len(self.RealSubject)):  #每个科
                printdata1=stdata.loc[:,['考号','班级','学号','姓名',self.RealSubject[RealSubjectid]+'考场号',self.RealSubject[RealSubjectid]+'座位号']]
                #stdata.to_excel(workdir+bag, sheet_name=self.RealSubject[RealSubjectid],index=False)
                
                #for j in range(len(printroomlist)): #要打印的考场号
                for index,roomid in enumerate(printroomlist): #要打印的考场号  
                    try:
                        worksheet = workbook.add_worksheet(str(self.RealSubject[RealSubjectid])+str(int(idroomlist[index])))#生成N个工作表 
                        #printdata=printdata1.loc[printdata1[self.RealSubject[RealSubjectid]+'考场号']-1==j]

                        printdata=printdata1.loc[printdata1[self.RealSubject[RealSubjectid]+'考场号']==int(idroomlist[index])]
                        printdata=printdata.copy()
                        printdata['缺考']=''
                        printdata=printdata.sort_values(by=[self.RealSubject[RealSubjectid]+'座位号'],na_position='first',ascending=(True))
                        self.myEntryBox
                        worksheet.merge_range(0,0,0,6,self.myEntryBox.get()+self.RealSubject[RealSubjectid]+'缺考登记表',title_format)        #表头
                        #worksheet.set_margins(left=0, right=0, top=0.8, bottom=0.8)   #页边距
                        rowheight=(a4height_V-100)/printdata.shape[0]
                        for i in range(printdata.shape[0]):
                            worksheet.set_row(i+2,rowheight)
                        worksheet.write_row(1, 0,  printdata.columns, cell_format)
                        worksheet.write_column(2, 0, printdata.iloc[:, 0], cell_format)
                        worksheet.write_column(2, 1, printdata.iloc[:, 1], cell_format)
                        worksheet.write_column(2, 2, printdata.iloc[:, 2], cell_format)
                        worksheet.write_column(2, 3, printdata.iloc[:, 3], cell_format)
                        worksheet.write_column(2, 4, printdata.iloc[:, 4], cell_format)
                        worksheet.write_column(2, 5, printdata.iloc[:, 5], cell_format)
                        worksheet.set_column(0,0,11)
                        
                        worksheet.conditional_format(1,0,printdata.shape[0]+1,6,{'type':'no_errors','format': cell_format})
                    except:
                        continue
                
            workbook.close()
            self.text.set(f'已保存到【{workdir}{bag}】文件!!!!')
            if self.output[1][0].get():printer_file(workdir+bag)   #打印开关打开就打印
            if self.output[0][0].get():xls2pdf.xlsx(self,bag)   #PDF开关打开就导出PDF 
                          
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo('出错',f'【{workdir}{bag}】文件已经打开？或'+workdir+'路径不存在？')   
               
    
    def kaowu(self):    #新，把各科组合一起，放入一张纸进行打印，太浪费了 2024/4/28 还是太浪费了，两个试室用一张A4纸
        workbook = xlsxwriter.Workbook(workdir+kaowu)
        roomdata=self.RoomData
        cell_title= workbook.add_format({
            'bold':     True,
            'border':   0,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':60,
            'text_wrap':1
        })
        cell_title_little_55= workbook.add_format({
            'bold':     True,
            'border':   0,
            'align':    'left',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':45,
            'text_wrap':1,
            'bottom_color': 'black',  # 底边框颜色
            'bottom': 1, 
        })
        cell_title_little_50= workbook.add_format({ #字数多的时候用这个，字号小
            'bold':     True,
            'border':   0,
            'align':    'left',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':35,
            'text_wrap':1,
            'bottom_color': 'black',  # 底边框颜色
            'bottom': 1, 
        })
        cell_title_with_border= workbook.add_format({
            'bold':     True,
            'border':   0,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':50,
            'text_wrap':1,
            'bottom': 2,        #设置下边框
        })
        
        printroomlist=list(range(int(self.startroom.get()),int(self.endroom.get())+1))   #要打印的考场号列表 
        testroomlist=self.RoomData.iloc[:,1].values.tolist()  #所有实际考场地点列表      
        idroomlist=self.RoomData.iloc[:,0].values.tolist()    #所有实际考场号列表 
        worksheet = workbook.add_worksheet('aaa')#生成N个工作表 
        startrow=0
        #colwidth=96         #列宽
        colwidth=a4width_H/2  #列宽
        firstheight=100     #第一行高#第一行高
        footheight=20       #空行行高
        tuningvalue=40      #微调值
        classid=-1              #用来做一个连续的计数
        startcol=0          #起始列
        for index,roomid in enumerate(printroomlist): #要打印的考场号  
            if int(roomid) not in idroomlist:continue
            classid+=1  
            testcounts=len(self.RealSubject)  
            worksheet.write(startrow,startcol,'试室'+str(int(idroomlist[classid])),cell_title)  #考室号  第一行
            worksheet.write(startrow+1,startcol,testroomlist[classid],cell_title_with_border)     #考室地址 第 二行  
            worksheet.set_row(startrow,firstheight)     #考室号  第一行
            worksheet.set_row(startrow+1,firstheight)   #考室号  第2行
            ##############  写入每个科目名及人数
            steprow=2
            writecontent=''
            for RealSubjectid in range(len(self.RealSubject)):
                writecontent=str(RealSubjectid+1)
                try:                    
                    subject=roomdata.loc[classid,self.RealSubject[RealSubjectid]]        #科目，信息1                    
                    roomStcount=int(roomdata.loc[classid,self.RealSubject[RealSubjectid]+'人数'])        #人数，信息2
                    writecontent=writecontent+'、'+subject+': '+str(roomStcount)
                    if len(writecontent)>=12:
                        worksheet.write(startrow+steprow,startcol,writecontent,cell_title_little_50)
                    else:
                        worksheet.write(startrow+steprow,startcol,writecontent,cell_title_little_55)     #考室地址 第三行  
                    steprow+=1
                except:
                    writecontent=writecontent+'、'
                    worksheet.write(startrow+steprow,startcol,writecontent,cell_title_little_55)     #考室地址 第三行  
                    steprow+=1
                    continue
            ##############计算各小行的行高
            subheight=(a4height_H-2*firstheight-footheight-tuningvalue)/testcounts      #每科占 行高
  
            for i in range(testcounts):
                worksheet.set_row(startrow+2+i,subheight)   #每科占 行高
            ############################
            worksheet.set_row(startrow+testcounts+3,footheight)   #空白行
            
            if classid%2==0:                 
                startcol=1
            else:
                startrow+=(3+testcounts)
                startcol=0       
            
        worksheet.center_horizontally()     #水平居中
        worksheet.set_column(0,1,colwidth)    #列宽
        #worksheet.set_portrait()        #竖排'''
        worksheet.set_landscape()
        worksheet.set_margins(left=0, right=0, top=0.4, bottom=0.6)   #页边距   
        try:
            workbook.close()
            self.text.set(f'已保存到【{workdir}{kaowu}】文件!!!!')
            if self.output[1][0].get():printer_file(workdir+kaowu)   #打印开关打开就打印  
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo('出错',f'【{workdir}{kaowu}】文件已经打开？或'+workdir+'路径不存在？')      
        if self.output[0][0].get():xls2pdf.xlsx(self,kaowu)   #PDF开关打开就导出PDF

    def inSideRoom(self):   #节约纸张打法
        workbook = xlsxwriter.Workbook(workdir+printroomin)
        roomdata=self.RoomData
        stdata=self.StData        
        diagonal_border = workbook.add_format({'diag_type': 1})
        cell_format = workbook.add_format({
            'bold':     False,
            'border':   1,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':10,
            'text_wrap':0 
        })
        num_format = workbook.add_format({  #考号太长，字体小一点
            'bold':     False,
            'border':   1,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':9,
            'text_wrap':0 
        })
        #testroomlist=self.RoomData.iloc[:,1].values.tolist()  #所有实际考场地点列表
        testroomlist=self.RoomData.iloc[:,1].values.tolist()[int(self.startroom.get())-1:int(self.endroom.get())]
        
        idroomlist=self.RoomData.iloc[:,0].values.tolist()    #所有实际考场号列表 
        printroomlist=list(range(int(self.startroom.get()),int(self.endroom.get())+1))   #要打印的考场号列表   [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
        #计算此考场要安排几场，3场可以横向，一行放2个。4-5场纵向。一行只能1个
        RealSubject=self.RealSubject        #考试科目 ['语数外', '物历', '化', '生', '政', '地']
        kaohaowidth=10.8    #一个科所占的列宽
        kaohaowidth_jieyu=9.63    #一个科所占的列宽 节约型
        
        roomNumwidth=3.5
        secondwidth=3.5
        everyheight=17        #有数据的行高
        blankheight=16        #空行行高
        blankwidth=1.5          #空列列宽
        if dpi==100:blankheight=18
        kaohaoheight=everyheight*3+blankheight   #三行数据，一个座位。所占的行高
        classid=0

        for index1,roomid in enumerate(printroomlist):
            if int(roomid) not in idroomlist:continue
            classid+=1  #用来做一个连续的计数
            try:
                worksheet = workbook.add_worksheet(str(int(roomid)))
                subSum=int((roomdata.iloc[classid-1].count()-2)/4)  #该考场的科目数量  =3  roomdata.iloc[index1]=该行数据
            except:
                continue             
            
            subFlag=0
            
            #width=subSum*kaohaowidth+(1.5+roomNumwidth+secondwidth)    #本座位所占的总宽度    42.1
            width=subSum*kaohaowidth_jieyu+(1.5+roomNumwidth+secondwidth) if subSum==4 else subSum*kaohaowidth+(1.5+roomNumwidth+secondwidth)

            '''if (a4width_H//width)*(a4height_H//kaohaoheight)>(a4width_V//width)*(a4height_V//kaohaoheight): 
                worksheet.set_landscape()   #横排可以排的学生数>竖排的学生数
                a4width=a4width_H
                cols=int(a4width//width)
                rows=int((a4height_H)//kaohaoheight)
                worksheet.set_margins(left=0, right=0, top=0.1, bottom=0.1)   #页边距
            else:
                worksheet.set_portrait()    #竖排
                a4width=a4width_V
                cols=int(a4width//round(width,0))   #2列
                rows=int(a4height_V//kaohaoheight)  #3，11                
                worksheet.set_margins(left=0, right=0, top=0.4, bottom=0.6)   #页边距
            #该考场为4科时，纵向排列，节省纸张
            if subSum==4:
                worksheet.set_portrait()    #竖排
                a4width=a4width_V
                cols=int(a4width//round(width,0))   #2列
                rows=int(a4height_V//kaohaoheight)  #3，11                
                worksheet.set_margins(left=0, right=0, top=0.4, bottom=0.6)   #页边距'''
            
            if (a4width_H//width)*(a4height_H//kaohaoheight)<=(a4width_V//width)*(a4height_V//kaohaoheight) or subSum==4: 
                worksheet.set_portrait()    #竖排
                a4width=a4width_V
                cols=int(a4width//round(width,0))   #2列
                rows=int(a4height_V//kaohaoheight)  #3，11                
                worksheet.set_margins(left=0, right=0, top=0.4, bottom=0.6)   #页边距
            else:
                worksheet.set_landscape()   #横排可以排的学生数>竖排的学生数
                a4width=a4width_H
                cols=int(a4width//width)
                rows=int((a4height_H)//kaohaoheight)
                worksheet.set_margins(left=0, right=0, top=0.1, bottom=0.1)   #页边距
            rowlist=[]
            for subjectid in range(len(RealSubject)): #获取同一试室中最大的行数
                try:
                    temp=int(roomdata[RealSubject[subjectid]+'人数'][classid-1])%cols
                    rows=int(roomdata[RealSubject[subjectid]+'人数'][classid-1])//cols+1 if temp else int(roomdata[RealSubject[subjectid]+'人数'][classid-1])//cols
                except:
                    continue
                rowlist.append(rows)        
            maxrow=max(rowlist)

            
            for subjectid in range(len(RealSubject)): #考试科目  RealSubject 有6科。但只有4科要考 RealSubject=['语数历物', '外语', '选修']
                seatNum=0
                subFlag+=1
                '''try:
                    checkvar=int(roomdata[RealSubject[subjectid]+'人数'][classid-1])
                except:
                    messagebox.showerror('错误',f'{RealSubject[subjectid]}科第{roomid}室人数有误，应为整数')
                    return'''
                try:
                    rows=int(roomdata[RealSubject[subjectid]+'人数'][classid-1])//cols+int(roomdata[RealSubject[subjectid]+'人数'][classid-1])%cols  #要打几行？52/2=26  #要打几行？52/2=26
                except:
                    subFlag-=1 #试室为空时，这里-1，后面的代码设置列时就不会往后跳了
                    print(f'注意：{RealSubject[subjectid]}{roomid}为空试室')
                    continue
                rows=int(roomdata[RealSubject[subjectid]+'人数'][classid-1])//cols+int(roomdata[RealSubject[subjectid]+'人数'][classid-1])%cols
                
                for col in range(cols):
                    worksheet.set_column(col*(3+subSum),col*(3+subSum), roomNumwidth)    #前两列列宽 
                    worksheet.set_column(col*(3+subSum)+1,col*(3+subSum)+1, secondwidth)    #前两列列宽 
                    #worksheet.set_column(col*(3+subSum)+2,col*(3+subSum)+subSum+1,kaohaowidth) #两个座位每科列宽 2 8 14    2 7 12
                    worksheet.set_column(col*(3+subSum)+2,col*(3+subSum)+subSum+1,kaohaowidth_jieyu) if subSum==4 else worksheet.set_column(col*(3+subSum)+2,col*(3+subSum)+subSum+1,kaohaowidth)
                    worksheet.set_column((1+col)*(3+subSum)-1,(1+col)*(3+subSum)-1,blankwidth) #两个座位之间的空格列宽 
                for row in range(rows):     #每一行。0-25
                    worksheet.set_row(row*4+1,everyheight)
                    worksheet.set_row(row*4+2,everyheight)
                    worksheet.set_row(row*4+3,everyheight)
                    #当尽量压缩空行行高时，最多一页放9行数据
                    worksheet.set_row(row*4,blankheight) if maxrow!=9 else worksheet.set_row(row*4,11.25)
                try:
                    tempdata=stdata[(stdata[RealSubject[subjectid]+'考场号']==roomid)]                    
                    tempdata=tempdata.sort_values(by=[RealSubject[subjectid]+'座位号'],na_position='first',ascending=(True)).reset_index() #按座位号排序
                    
                except:
                    messagebox.showerror('错误','学生信息表头有误，未找到场室信息，请检查~')
                    return
                for j in range(maxrow):   #为了响应要求，在空白的座位处也要写上科目字样，让它全部循环完
                    for i in range(cols): #行循环  2                        
                        seatNum+=1
                        #if seatNum>int(roomdata[RealSubject[subjectid]+'人数'][classid-1]):break  #循环超过实际学生数，下面不进行
                        worksheet.conditional_format(4*j+1,i*(3+subSum),4*j+3,i*(3+subSum)+1+subFlag, {'type': 'blanks','format': cell_format})
                        #为了让每一页都打印一样，应该在第一页上面加一个空行，这样后面就和它一样了
                        worksheet.write(4*j+1,i*(3+subSum),'试室',cell_format)
                        worksheet.write(4*j+2,i*(3+subSum),'地址',cell_format)
                        worksheet.write(4*j+3,i*(3+subSum),'座位',cell_format)
                        worksheet.write(4*j+1,i*(3+subSum)+1,roomid,cell_format)
                        try:
                            worksheet.write(4*j+2,i*(3+subSum)+1,testroomlist[classid-1],num_format) #地址字号小一点
                            worksheet.write(4*j+3,i*(3+subSum)+1,seatNum,cell_format)                            
                            worksheet.write(4*j+1,i*(3+subSum)+1+subFlag,tempdata.at[seatNum-1,'考号'],num_format) #地址字号小一点
                            worksheet.write(4*j+2,i*(3+subSum)+1+subFlag,RealSubject[subjectid],cell_format)
                            worksheet.write(4*j+3,i*(3+subSum)+1+subFlag,tempdata.at[seatNum-1,'姓名'],cell_format)
                        except: 
                            #seatNum-=1  #当时是为了什么要-1，忘了，，，，
                            worksheet.write(4*j+1,i*(3+subSum)+1+subFlag,'',diagonal_border)
                            worksheet.write(4*j+2,i*(3+subSum)+1+subFlag,RealSubject[subjectid],cell_format)
                            worksheet.write(4*j+3,i*(3+subSum)+1+subFlag,'',diagonal_border)
                            continue
  
            worksheet.center_horizontally()     #水平居中
            worksheet.center_vertically()       #垂直居中
            worksheet.set_header(header='',options=None,margin=0)
            worksheet.set_paper(9)      #9=A4纸    
        try:
            workbook.close()

            if self.output[0][0].get():xls2pdf.xlsx(self,printroomin)   #PDF开关打开就导出PDF
            self.text.set(f'已保存到【{workdir}{printroomin}】文件!!!!')
            if self.output[1][0].get():printer_file(workdir+printroomin)   #打印开关打开就打印 
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo('出错',f'【{workdir}{printroomin}】文件已经打开？或'+workdir+'路径不存在？')      
    
    def outSideRoom(self):########考场外粘贴表
        testroomlist=self.RoomData.iloc[:,1].values.tolist()  #所有实际考场地址列表        
        idroomlist=self.RoomData.iloc[:,0].values.tolist()    #所有实际考场号列表     
        printroomlist=list(range(int(self.startroom.get()),int(self.endroom.get())+1))   #要打印的考场列表        
        workbook = xlsxwriter.Workbook(workdir+printroomout) 
        addcol=3      #2023/12/21改个功能，以前加人都是第1列，想换到手动指定的列，这是从第3列开始加人
        merge_format = workbook.add_format({
            'bold':     False,
            'border':   0,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':11,
            'text_wrap':1 
        })
        title_format = workbook.add_format({
            'bold':     True,
            'border':   3,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':15,
            'text_wrap':0 
        })
        cell_format = workbook.add_format({
            'border':   1,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':10,
            'text_wrap':0 
        })
        temp=[]
        roomaddpd=pd.DataFrame()
        c=[x*4+3 for x in range(len(self.RealSubject))] #建立列表，储存人数所在的列[3,7,11 15 19 23]      self.RealSubject=['语数外', '物历', '化', '生', '政', '地']  

        for i in range(len(self.RealSubject)):
            roomaddlist=[]      #把每个场子的人数差放入列表
            temp.append((i+1)*4)
            a=b=0
            for j in printroomlist:     #printroomlist=[1,2,3,-----20]
                try:                
                    a=self.RoomData.loc[j-1,self.RealSubject[i]+'人数'] #本场实际人数   
                    b=self.RoomData.loc[j-1,self.RealSubject[i]+'行']*self.RoomData.loc[j-1,self.RealSubject[i]+'列'] #标称行*列
                    maxcount=int(listmax(self.RoomData.iloc[j-1,c].dropna().tolist()))    #人数最大的科                    
                except:
                    continue
                if a==maxcount:
                    chazhi=a-b if (a-b)>0 else 0    #b是安排的行列数相乘。a是实际人数，如果实际人数多，则要加行
                    roomaddlist.append(chazhi)
                else:
                    roomaddlist.append(0)                
            roomaddpdtemp=pd.DataFrame(roomaddlist)
            roomaddpd=pd.concat([roomaddpd,roomaddpdtemp],axis=1) #得到一个差值表，储存各室的相差人数，如有正数说明要加人
        
        rowNumdata=self.RoomData.iloc[:,temp] #一个series，放置所有试室的打印行       

        for index,subject in enumerate(self.RealSubject):            #每个科目都要问一次
            if (messagebox.askyesno('提示', '需要【生成】'+ subject +'的考场外张贴表吗？')):                
                room_col_list=self.RoomData[subject+'列'][self.RoomData[subject+'列'].notnull()].values.tolist()  #考场列列表
                roomcount=self.RoomData[subject][self.RoomData[subject].notnull()].values.tolist()
                colcount=0      #计数，用来为每个科的列数计数
                if int(self.endroom.get())>len(roomcount):
                    printroomlist=list(range(int(self.startroom.get()),int(self.endroom.get())+1))  #要打印的考场列表
                filtered_list = self.RoomData[self.RoomData[subject+'人数'].notnull()].index.tolist()       #找到考场安排表中有数据的行

                for index1,roomid in enumerate(printroomlist):  #在名单表中找该考场的所有行 [3,10]
                    
                    if roomid-1 in filtered_list:
                        room_st_count=0
                        colcount+=1
                        
                        printdata=self.StData[self.StData[subject+'地址'].isin([testroomlist[roomid-1]])]
                        printdata=printdata.sort_values(by=[subject+'座位号'],na_position='first',ascending=(True)) #按座位号排序
                        worksheet = workbook.add_worksheet(subject+str(int(idroomlist[roomid-1])))
                        worksheet.merge_range('A1:B1','试室地址：',merge_format)  #
                        worksheet.merge_range('C1:D1',testroomlist[roomid-1],merge_format) 
                        worksheet.merge_range('A2:B2','试室号：',merge_format)  #
                        worksheet.merge_range('C2:D2',idroomlist[roomid-1],merge_format)   
                        worksheet.merge_range('E1:F2',subject,merge_format) #科目，直接用表头算了
                        worksheet.merge_range('G1:I2','讲    台',title_format) 

                        worksheet.merge_range(0,9,2,int(room_col_list[0])*3-2,self.myEntryBox.get(),merge_format) #考试名称
                        ###########设置页面#################
                        worksheet.center_horizontally()     #水平居中
                        worksheet.center_vertically()       #垂直居中
                        worksheet.set_margins(left=0.3, right=0.3, top=0.1, bottom=0.17)   #页边距
                        worksheet.set_header(header='',options=None,margin=0)
                        worksheet.set_paper(9)      #9=A4纸            
                        worksheet.set_landscape()   #纸张方向横向
                        ##################################
                        newdata=printdata.drop(index=printdata.index)   #原表清空做一个空表
                        room_st_num=printdata.shape[0]
                        restST=int(listmax(roomaddpd.iloc[index1].values.tolist()))
                        restST1=restST
                        for col in range(int(room_col_list[colcount-1])):   #每个室共几列？一列一列写room_row_list=[7,7,7,6,5]
                            if restST>0 and (addcol-1)<=col:    #多人在这里处理，可否加一个条件，第几列加人，加一个手动开关
                                addst=1
                                restST-=1
                            else:
                                addst=0
                            temp1=listmax(rowNumdata.iloc[index1].dropna().values.tolist())

                            maxrow=int(temp1)
                            newdata=printdata.iloc[:maxrow+addst,:] #取数据表前N行。本来是用前面的，现在要取的数据应是最多的那个科目的行数。
                            newdataindex=printdata.iloc[:maxrow+addst,:].index    #索引
                            printdata.drop(labels=newdataindex,inplace=True)    #原数据中删除已排的index
                            #如果是偶数列，newdata要倒序排列
                            if col%2==1:newdata=newdata.sort_values(by=[subject+'座位号'],na_position='first',ascending=(False))
                            for row in range(maxrow+addst):#这里应是几科里最大的行数。
                                room_st_count+=1
                                if room_st_count>room_st_num:break      #大于实际人数了，终止循环
                                try:
                                    kaohao=newdata['考号'].tolist()[row]
                                except:
                                    print('出错，跳过')
                                    continue
                                worksheet.write(5*row+3,3*col,'考号',cell_format)
                                worksheet.write(5*row+4,3*col,'姓名',cell_format)
                                worksheet.write(5*row+5,3*col,'试室号',cell_format)
                                worksheet.write(5*row+6,3*col,'座位号',cell_format)
                                worksheet.write(5*row+3,3*col+1,kaohao,cell_format)
                                worksheet.write(5*row+4,3*col+1,newdata['姓名'].tolist()[row],cell_format)
                                worksheet.write(5*row+5,3*col+1,newdata[subject+'考场号'].tolist()[row],cell_format)
                                worksheet.write(5*row+6,3*col+1,newdata[subject+'座位号'].tolist()[row],cell_format)
                                
                            worksheet.set_column(3*col,3*col, 5)    #考号列列宽   
                            worksheet.set_column(3*col+1,3*col+1, 11)    #考号列列宽
                            worksheet.set_column(3*col+2,3*col+2, 1.5)    #空列列宽
                        addrow=int(restST1/(room_col_list[colcount-1])+0.99) #多加的行数
                        maxrows=maxrow+addrow
                        blankrow=12 #空行高12
                        rowsheight=(a4height_H-(maxrows-1)*blankrow)/(maxrows*4)
                        if rowsheight<10:rowsheight=10   #限制行高
                        if rowsheight>21:rowsheight=20
                        ##print(rowsheight,123456,roomid)
                        for row in range(10): #设置行高                        
                            worksheet.set_row(5*row+7,blankrow)   #设置空行行高
                            worksheet.set_row(5*row+3,rowsheight)   #设置数据行行高
                            worksheet.set_row(5*row+4,rowsheight)   #设置数据行行高
                            worksheet.set_row(5*row+5,rowsheight)   #设置数据行行高
                            worksheet.set_row(5*row+6,rowsheight)   #设置数据行行高  
                #roomaddid+=1                       
        
        try:            
            workbook.close()            
            self.text.set(f'已保存到【{workdir}{printroomout}】文件!!!!')
            if self.output[0][0].get():xls2pdf.xlsx(self,printroomout)   #PDF开关打开就导出PDF
            if self.output[1][0].get():printer_file(workdir+printroomout)   #打印开关打开就打印  

        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo('出错',f'【{workdir}{printroomout}】文件已经打开？或'+workdir+'路径不存在？')        

    def ClassTable(self):
        #班级核对表
        workbook = xlsxwriter.Workbook(workdir+printclass) 
        merge_format = workbook.add_format({
            'bold':     True,
            'border':   0,
            'align':    'center',#水平居中
            'valign':   'vcenter',#垂直居中
            'font_name':u'微软雅黑',
            'font_size':15
        })
        kaohao_width_base=11.5
        class_width_base=2.5
        name_width_base=6
        subject_width_base=5.6

        subject=len(self.RealSubject)*3   #因为一科占用了3列，原来核对表里1科只占2列，*2，现在3列，*3
        endcol=len(anpai.delStrFromList(self,list(self.StData)))+1 #ST表列标签最后一列        
        startcol=endcol-subject     #需要的几列数据
        startclass=int(self.startclass.get())  #要打印的起始班 从entry框中得到
        endclass=int(self.endclass.get())
        a4height=a4height_V-20      #高度微调
        a4width=a4width_V
        format_border = workbook.add_format({'border': 1})
        font_format=workbook.add_format({'font_name': u'微软雅黑','font_size':10,'align':'center','valign':'vcenter','text_wrap':1 })  #文字格式
        indexdata=self.StData.sort_values(by=['班级','学号'],na_position='first',ascending=(True,True)).reset_index()
        '''上面一行，原来是直接对self.stdata进行了更改，多加了一列index,reset_index会增加新列。导致多运行几次就出错 ，
        所以不要对原self改动，用新变量就好了'''
        for i in range(startclass,endclass+1):
            try:
                rowlist=indexdata[indexdata.班级==i].index.tolist()
                startRow=rowlist[0]
                endRow=rowlist[-1]+1               
                
                printData=pd.concat([indexdata.iloc[startRow:endRow,1:5],indexdata.iloc[startRow:endRow,startcol:endcol]],axis=1)  #需要打印的数据 ，前4列，后几列
                
                for j in range(len(self.RealSubject)):
                    printData.drop(columns=[self.RealSubject[j]+'考场号'],axis=1,inplace=True)  #班级核对表不用考场号 删除   
                printData = printData.dropna(axis=1, how='all')     #当某列数据全为nan时，删除此列
                worksheet = workbook.add_worksheet(str(i))      #增加一个工作表。以班名为工作表名
                worksheet.write_row(1,0,list(printData),font_format) #写表头            
                worksheet.set_row(0,40)         #第一行 标题，行高多一点
                worksheet.set_row(1,28)         #第一行 标题，行高多一点
                for j in range(len(list(printData))): #一列一列地写
                    worksheet.write_column(2,j,delnanfromlist(list(printData.iloc[:,j])))  #write_row write_column可以直接写一个列表，方便写数据
                    #这里为什么要一列列写，不能直接写全部的data吗？下次看看
            except:
                continue        
            #下面设置工作表格式
            worksheet.center_horizontally()     #水平居中
            worksheet.set_margins(left=0, right=0, top=0.1, bottom=0)   #页边距
            worksheet.merge_range(0,0,0,len(list(printData))-1,self.myEntryBox.get()+str(i)+'班座位表',merge_format)  #标题文字
            worksheet.set_header(header='',options=None,margin=0)
            worksheet.set_paper(9)      #9=A4纸
            format_border =  workbook.add_format({'bold': False, 'font_color': 'black','font_size':14,'border': 1})   # 设置边框格式
            worksheet.conditional_format(1,0,printData.shape[0]+1,printData.shape[1]-1, {'type':'no_errors','format': format_border})
            
            sub_count=printData.shape[1]-4      #有几科
            rate=(kaohao_width_base+class_width_base*2+name_width_base+sub_count*subject_width_base)/a4width      #缩放比例
            worksheet.set_column(0, 0, kaohao_width_base/rate)    #考号列列宽
            worksheet.set_column(1, 2, class_width_base/rate)    #班级学号列列宽
            worksheet.set_column(3, 3,name_width_base/rate)    #姓名列列宽
            worksheet.set_column(4, printData.shape[1]-1, subject_width_base/rate) #后面列的列宽

            everyheight=self.queryheight(a4height,printData.shape[0])
            for i in range(2,printData.shape[0]+2):     #所有行的行高                
                worksheet.set_row(i,everyheight,font_format)   #动态分布行高
        try:
            workbook.close()
            self.text.set(f'已保存到【{workdir}{printclass}】文件!!!!')
            if self.output[0][0].get():xls2pdf.xlsx(self,printclass)   #PDF开关打开就导出PDF
            if self.output[1][0].get():printer_file(workdir+printclass)   #打印开关打开就打印  
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo('出错','【班级核对表.xlsx】文件已经打开！！关闭文件再来一次~~')   
    def queryheight(self,height,datasum):
        everyheight=round(height/datasum,1)
        while True:
            if everyheight*datasum>height:
                everyheight=everyheight-0.1
            else:
                if everyheight>14.9:everyheight=14.9
                return everyheight

    def exit2(self):
        global flagWin2
        flagWin2=False
        self.destroy()
    def open(self,event=None):
        global workdir
        self.printFile =filedialog.askopenfilename(filetypes=[('Excel','.xls .xlsx .xlsm')])
        self.openfile1.delete(0,"end")
        self.openfile1.insert(0,self.printFile)
        testname=re.findall(r'[^\\_:*?"<>|^\\/:*?"<>]+$',self.printFile)[0] #找/号或_号右边的字符串，就是文件名了。。
        testname=re.findall(r'(.+?)\.',testname)[0]
        
        self.myEntryBox.delete(0,'end')
        self.myEntryBox.insert(0,testname)
        
        temp=re.findall(r'[^\\/:*?"<>|\r\n]+$',printFileUrl)[0]
        temp1=re.findall(r'[^\\/:*?"<>|\r\n]+$',self.printFile)[0]
        try:
            workdir=re.findall(r'(^.+\/).',self.printFile)[0]
        except:
            workdir=''

    def cancel(self):
        self.Retcode = None
        self.exit2()

class xls2pdf:
    def __init__(self, pathname, export='.'):
        self._handle_postfix = ['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']
        self._filename_list = list()
        self._export_folder = os.path.join(os.path.abspath('.'), 'pdfconver')
        if not os.path.exists(self._export_folder):
                os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)
    
    def _enumerate_filename(self, pathname):
        '''
        读取所有文件名
        '''
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError('文件 {} 后缀名不合法！仅支持如下文件类型：{}。'.format(pathname, '、'.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError('文件/文件夹 {} 不存在或不合法！'.format(pathname))

    def _is_legal_postfix(self, filename):
        return filename.split('.')[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith('~')
    
    def run_conver(self):
        '''
        进行批量处理，根据后缀名调用函数执行转换
        '''
        print('需要转换的文件数：', len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split('.')[-1].lower()
            funcCall = getattr(self, postfix)
            print('原文件：', filename)
            funcCall(filename)
        print('转换完成！')
    
    def doc(self, filename):
        '''
        doc 和 docx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        print('保存 PDF 文件：', exportfile)
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        w = Dispatch("Word.Application")
        doc = w.Documents.Open(filename)
        doc.ExportAsFixedFormat(exportfile, constants.wdExportFormatPDF,
                Item=constants.wdExportDocumentWithMarkup,
                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)        
        w.Quit(constants.wdDoNotSaveChanges)
    
    def docx(self, filename):
        self.doc(filename)
    
    def xlsx(self, filename):
        filename=workdir+filename
        exportfile = filename.split('.')[0] + '.pdf'
        try:
            xlApp = DispatchEx("Excel.Application")
            xlApp.Visible = False    
            xlApp.DisplayAlerts = 0   
            
            books = xlApp.Workbooks.Open(filename,False)
            books.ExportAsFixedFormat(0, exportfile)
            books.Close(False)
        except:
            messagebox.showerror('错误','错误了，退出再试试，是不是excel文件没正常关闭？')
        messagebox.showinfo('成功',f'已保存到{exportfile}，直接在windows里自己打印吧~~')
        xlApp.Quit()
class AppMain(tk.Tk):
    def __init__(self):
        super().__init__()
        style = ttk.Style()
        style.configure('danger.TButton', font='Times 10', foreground='red', padding=0,background='white', border=0)
        style.configure('blue.TButton', font='Times 10', background='blue',padding=0,foreground='blue', border=0)
        style.configure('orange.TButton', font='Times 10', background='orange',padding=0,foreground='orange', border=0)
        style.configure('green.TButton', font='Times 10', background='green',padding=0,foreground='green', border=0)
        global flagWin1,flagWin2,flagWin3,flagWin4
        flagWin1=flagWin2=flagWin3=flagWin4=False
        self.title("考场安排 By longlong")
        rootX=self.winfo_screenwidth()
        rootY=self.winfo_screenheight()
        #self.wm_attributes('-topmost',True)     #root窗口置顶
        self.geometry("%dx%d+%d+%d" %(x,y,(rootX-x)/2,(rootY-y)/2))
        self.protocol("WM_DELETE_WINDOW", self.exitMainApp) #点击关闭按钮触发
        self.setup_UI()
        self.after(500, self.start_check_update_thread) #窗口加载完成后，延迟 500 毫秒启动检查更新
    def setup_UI(self):        
        if not license():
            sys.exit(1)
        self.fuction1 = ttk.Button(self, text='考场安排',command=self.func1,bootstyle="success").place(x=100,y=100, width=200, height=40)
        self.fuction2 = ttk.Button(self, text='打印座位表', command=self.print,bootstyle="success").place(x=100,y=170, width=200, height=40)    
        self.closeButton=ttk.Button(self, text='关闭窗口', command=self.exitMainApp,bootstyle="success").place(x=100,y=240, width=200, height=40)        
    def start_check_update_thread(self):
        """开启后台线程去检查更新，防止卡顿界面"""
        threading.Thread(target=self.check_for_updates, daemon=True).start()
    def get_remote_version(self, raw_url):
        """通过 Raw 链接获取远程脚本内容，并用正则提取版本号"""    
        try:
            response = requests.get(raw_url, timeout=5)
            if response.status_code == 200:
                # 获取二进制内容并解码为字符串用于正则匹配
                remote_bytes = response.content
                remote_code_str = remote_bytes.decode("utf-8")            
                
                # 兼容双引号和单引号
                match = re.search(r'__version__\s*=\s*[\'"]([^\'"]+)[\'"]', remote_code_str)            
                if match:
                    remote_version = match.group(1)
                    # 返回版本号字符串，以及【二进制字节流】用于直接安全写入
                    return remote_version, remote_bytes
                else:
                    print("在远程脚本中未找到 __version__ 变量")
                    return None, None
            else:
                print(f"获取远程脚本失败，状态码: {response.status_code}")
                return None, None
        except Exception as e:
            print(f"请求网络出错: {e}")
            return None, None

    def update_self(self, remote_bytes, remote_ver):
        """直接用新代码的二进制字节流覆盖本地当前脚本文件"""
        try:
            # 获取当前运行脚本的绝对路径
            current_file = os.path.abspath(sys.argv[0])        
            with open(current_file, "wb") as f:
                f.write(remote_bytes)          
            messagebox.showinfo('更新成功', f'成功更新到版本 {remote_ver}！请重新启动程序以应用更改。')
            print("🎉 更新成功！请重新启动程序以应用更改。")
            sys.exit(0) # 退出当前旧进程
        except Exception as e:
            print(f"❌ 写入文件失败（可能由于文件权限或正被占用）: {e}")
    def parse_version(self,version_str):
        """将版本号字符串转换为整数元组，例如 '1.10.2' -> (1, 10, 2)"""
        version_str = version_str.lstrip("v")
        try:
            return tuple(map(int, version_str.split(".")))
        except ValueError:
            return (0,)
    def check_for_updates(self): #检查远程脚本并更新
        print('开始更新')
        raw_url = f"http://tv.skyyareply.eu.org/https://raw.githubusercontent.com/skyyaman/tvbox/refs/heads/main/python/examroom.py"
        remote_ver, remote_code = self.get_remote_version(raw_url)
        if not remote_ver:
            return    
        print(f"当前版本: {__version__}")
        print(f"远程版本: {remote_ver}")
        
        # 对比版本大小
        if self.parse_version(remote_ver) > self.parse_version(__version__):
            print(f"发现新版本 ({remote_ver})！")
            self.update_self(remote_code,remote_ver)
        else:
            print("当前已是最新版本，无需更新。")
    def setup_config2(self):
        global flagWin1,flagWin2
        if flagWin1:
            self.Input1.focus()
            return
        if flagWin2:
            self.Print1.focus()
            return
        res = self.print_voucher()
        if res is None: return

    def func1(self):
        '''对应Win1Input窗口类'''
        self.Input1 = anpai()
        self.wait_window(self.Input1)
        try:
            return self.Input1.Retcode
        except:
            pass
    
    def print(self):
        self.Input1 = Win2Print()
        self.wait_window(self.Input1)
        try:
            return self.Input1.Retcode
        except:
            pass
    
    def exitMainApp(self):     
        self.destroy()
def license(): #简单授权
    try:
        url='https://gitee.com/xtzx/tv/raw/master/1.txt'
        key=requests.get(url=url,timeout=(3.05, 6.06))
        return True if key.text.strip().replace('\\n','')=='longlong is your daddy' else  False
    except:
        messagebox.showinfo('出错','不能联网，请确认网络是否正常再试一次抬抬/by longlong？')
        return False
def getname(list1):
    colName=[]
    for i in range(int(len(list1)/4)):colName.append(list1[4*i+2])  #得到科目列表
    return colName

def listmax(list):#用来在list中去掉文本，并取list中的最大值
    for i in list:
        try:
            if '+' in i:
                list.remove(i)
                continue
            if i.isalnum():list.remove(i)
        except:
            continue
    max1=max(list)
    return max1

def delnanfromlist(mylist):         #从列表中替换nan为空
    for i in range(len(mylist)):
        if pd.isnull(mylist[i]):
            mylist[i]=''
    return mylist

def alter(file,old_str,new_str):
    file_data = ""
    replaced=0
    with open(file, "r", encoding="utf-8") as f:
        for line in f:
            if old_str in line:
                replaced+=1
                if replaced==1: line = line.replace(old_str,new_str)
            file_data += line
    with open(file,"w",encoding="utf-8") as f:
        f.write(file_data)
def printer_file(filename):
    app = xw.App(visible=False,add_book=False) #启动Excel程序
    workbook = app.books.open(filename) #打开要打印的工作簿
    worksheet = workbook.sheets  #所有工作表
    
    for i in worksheet:
        i.api.PrintOut(Copies=1,Collate=True) #打印指定工作表
    workbook.close()
    app.quit()
    print(filename+'----打印成功')
def getdpi():
    w = GetSystemMetrics (0)
    real_w = win32print.GetDeviceCaps(win32gui.GetDC(0), win32con.DESKTOPHORZRES)
    screen_scale_rate = round(real_w / w, 2)
    screen_scale_rate = int(screen_scale_rate * 100)
    return screen_scale_rate

if __name__ == '__main__':    
    x=390
    y=380
    dpi=getdpi()

    if dpi==100:    
        a4height_V=758  #竖排时宽高  850---768
        a4width_V=95    #           90
        a4height_H=510  #横排时高   550-440
        a4width_H=126   #横排时宽           123---100
    else:#dpi=125
        a4height_V=830  #竖排时宽高  850---768
        a4width_V=93    #           93----93
        a4height_H=530  #横排时宽高   550-440
        a4width_H=123   #           123---100
    __DIR__ = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))  #程序所在的目录
    workdir=''
    examname='考试'
    fileurl=os.path.join(os.path.expanduser("~"), 'Desktop')
    fileurl=u'点击右边按钮选择需要安排的考场数据'
    printFileUrl=u'点击右边按钮选择已安排数据'
    printroomout=u'temp_考场室外张贴表.xlsx'
    printroomin=u'temp_考场课桌张贴表.xlsx'
    printclass=u'temp_班级核对表.xlsx'
    bag=u'temp_试卷袋打印.xlsx'
    kaowu=u'temp_考务室标签.xlsx'
    app = AppMain()    
    app.mainloop()