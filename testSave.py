import pandas as pd
import pickle, sys


class jiankao:

    # IT老师可监考人
    def teachers(self, teach):
        try:

            print('这是本次参与的监考人：', teach)

            choice = True
            while choice:
                changeT = input('是否需要添加or删除，添加输入1，删除输入2，不修改输入0：')

                if (changeT == '1'):
                    addT = input('输入需要添加的监考人名： ')
                    teach.append(addT)

                elif (changeT == '2'):
                    dT = input('输入需要删除的监考人名： ')
                    teach.remove(dT)
                else:
                    choice = False
                    print('本次监考确定监考人：', teach)
                    return teach
        except:
            print('哎哟喂，程序‘IT老师可监考人’出现错误了，要不你试一试重新运行或者来找我')

    # 查阅监考老师课程表
    def searchT(self):
        try:
            teacher = ['林颖', '曹明岩', '刘畅', '王媛']


            print('请注意：程序不支持不接受不分析表格的横向合并，只支持接受竖向合并！！')

            schedule = input('输入课程表目录(如 F:\Python_Office\总课表.xlsx)：')

            df = pd.read_excel(schedule)

            df.fillna(method='ffill')

            day = input('请输入考试日期，如 星期一：')

            listt = self.teachers(teacher)

            key_c = [col for col in df.columns if day in col]

            # 储存一二节，三四，五六，七八
            yier = []
            ye = []
            sansi = []
            ss = []
            wuliu = []
            wl = []
            qiba = []
            qb = []

            for t in listt:
                for i in key_c:
                    a = df.loc[df[i].str.contains(t, na=False), i]
                    for l, c in a.items():
                        if '.1' in a.name:
                           # print(day + ' 三四节：\n', c)
                            sansi.append(c)
                        elif '.2' in a.name:
                            # print(day+ ' 五六节：\n', c)
                            wuliu.append(c)
                        elif '.3' in a.name:
                            # print(day +' 七八节：\n', c)
                            qiba.append(c)
                        else:
                            # print(day +' 一二节：\n', c)
                            yier.append(c)

            self.addlist(ye, listt, yier)
            print(day,'的一二有课： ')
            print(ye)

            self.addlist(ss, listt, sansi)
            print(day, '的三四有课： ')
            print(ss)

            self.addlist(wl, listt, wuliu)
            print(day, '的五六有课： ')
            print(wl)

            self.addlist(qb, listt, qiba)
            print(day, '的七八有课： ')
            print(qb)

            self.searchNotT(key_c, listt, df)

        except:
            print('哎哟喂，程序‘查阅监考老师课程表’出现错误了，要不你试一试重新运行或者来找我')

    def addlist(self, tt, lista, sansi):
        for i in lista:
            for t in sansi:
                if i in t:
                    tt.append(i)
        return tt

    # 查阅没有课的监考老师课程表
    def searchNotT(self, key_c, listt, df):
        try:
            OutT = []

            for t in listt:
                for i in key_c:
                    a = df.loc[df[i].str.contains(t, na=False), i]
                    for l, c in a.items():
                        OutT.append(c)
            name = []
            for i in listt:
                for n in OutT:
                    if i in n:
                        name.append(i)
            name = sorted(set(name), key=name.index)

            for tea in listt:
                if tea not in name:
                    print('今天没有课的老师有： ', tea)

        except:
            print('哎哟喂，程序‘查阅没有课的监考老师课程表’出现错误了，要不你试一试重新运行或者来找我')

    # 计算监考次数
    def count(self):
        try:
            while True:
                load = input('输入1：统计监考人员； 输入2：查询次数；输入其他返回上一层；请输入：')
                if load == '1':
                    print('注意：监考人员格式请竖着排列')
                    teacher_list = input('输入本次监考目录(如F:\Python_Office\总课表.xlsx)：')
                    td = pd.read_excel(teacher_list, header=None)
                    count_f = []
                    file_load = self.load_file('countZero.pickle', count_f)
                    # fl=open('countZero.pickle','rb')
                    # file_load=pickle.load(fl)
                    df_v = td.values.T[0].tolist()
                    # for k, v in file_load.items():
                    #   for i in df_v:
                    #      if i == k:
                    #         v = v + 1
                    #        file_load[k] = v

                    # d=self.sortTimes(file_load)
                    # print(d)

                    # f = open('countZero.pickle', 'ab')
                    # pickle.dump(file_load, f)
                    # f.close()
                    for i in count_f:  # 增加新可统计老师
                        for j in i:
                            for a in df_v:
                                if j == a:
                                    c = i[a] + 1
                                    i[a] = c
                                    # print(a,c)
                    for i in count_f:
                        print(i)
                        f = open('countZero.pickle', 'ab')
                        pickle.dump(i, f)
                        f.close()

                elif load == '2':
                    with open('countZero.pickle', 'rb') as f:
                        while True:
                            try:
                                r = pickle.load(f)
                                print(r)
                            except EOFError:
                                break
                    # fl = open('countZero.pickle', 'rb')
                    # file= pickle.load(fl)
                    # d=self.sortTimes(file)
                    # print(d)
                    # fl.close()

                else:
                    break
        except:
            print('哎哟喂，程序‘计算监考次数’出现错误了，要不你试一试重新运行或者来找我')

    def save_file(self, v, fname):
        f = open(fname, 'wb')
        pickle.dump(v, f)
        f.close()

    def load_file(self, fname, count_f):
        with open(fname, 'rb') as f:
            while True:
                try:
                    r = pickle.load(f)
                    count_f.append(r)
                    # print(r)
                except EOFError:
                    break

    # 监考次数排序，从小到大
    def sortTimes(self, a):
        l = sorted(a.items(), key=lambda d: d[1])
        return l

    # 初始化监考次数为0
    def new(self):
        try:
            countT = {'林颖': 0, '曹明岩': 0, '李发金': 0}
            while True:
                choice = input('输入 1 重置次数为零；输入2，增加新老师进入统计；其他为返回')
                if choice == '1':
                    self.save_file(countT, 'countZero.pickle')
                    # f = open('countZero.pickle', 'wb')
                    # pickle.dump(countT, f)
                    # f.close()
                    print('所有老师监考次数以重置为0！！！')
                    # 增加新的可统计老师
                elif choice == '2':
                    new_name = input('新的统计老师名字：')
                    countT[new_name] = 0
                    f = open('countZero.pickle', 'ab')
                    pickle.dump({new_name: 0}, f)
                    f.close()
                    print('已加入新成员')
                else:
                    break
        except:
            print('哎哟喂，程序‘初始化监考次数’出现错误了')

    def main(self):

        print('本程序目前只支持excel 文档！！！')
        print('第一次用的时候请首先输入0，重置监考次数全为零')
        while True:
            check = input('输入 1 查询课表；输入 2 查询并统计监考次数；重置监考次数和添加新监考老师输入0；输入其他按键为退出程序： ')
            if check == '1':
                self.searchT()

            elif check == '2':
                self.count()

            elif check == '0':
                self.new()

            else:
                print('程序退出')
                sys.exit()
                break


jiankao1 = jiankao()
# jiankao1.new()
# jiankao1.searchT()
jiankao1.main()
