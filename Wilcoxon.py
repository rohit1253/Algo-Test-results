class wilcoxon():
    abs_diff = []
    sign = []
    cl_no_zero_diff = []
    rank = []
    sorted_abs_diff = []
    freq = []
    cumfreq = []
    wil_ranks = []
    var = 0
    total = 0
    pos_sum = 0
    neg_sum = 0
    
    def __init__(self, fname, sh_index, x, y):
        '''fname is the filename, sh_index is the sheet index of the workbook,
        x is the key product column index and y is the non-key product column index'''
        self.fname = fname
        self.sh_index = sh_index
        self.x = x
        self.y = y
        
    def read_excel(self):
        '''Reads a sheet from a excel workbook'''
        import xlrd
        wb = xlrd.open_workbook(self.fname)
        sheet = wb.sheet_by_index(self.sh_index)
        return sheet
    
    def wilcox_rank(self, sheet):
        
        for i in range(1, sheet.nrows):
            wilcoxon.abs_diff.append(abs(sheet.cell_value(i, self.x) - sheet.cell_value(i, self.y)))
            if (sheet.cell_value(i, self.x) - sheet.cell_value(i, self.y))<0:
                wilcoxon.sign.append(-1)
            else:
                wilcoxon.sign.append(1)
        comb_list = sorted(list(zip(wilcoxon.sign, wilcoxon.abs_diff)),key=lambda x: x[1])
        for i in range(len(comb_list)):
            if comb_list[i][1] != 0.0:
                wilcoxon.cl_no_zero_diff.append(comb_list[i])
        for i in range(len(wilcoxon.cl_no_zero_diff)):
            wilcoxon.rank.append(i+1)
            wilcoxon.sorted_abs_diff.append(wilcoxon.cl_no_zero_diff[i][1])
        d = {x:wilcoxon.sorted_abs_diff.count(x) for x in wilcoxon.sorted_abs_diff}
        for i in d.values():
            wilcoxon.freq.append(i)
        for i in range(len(wilcoxon.freq)):
            wilcoxon.var = wilcoxon.var + wilcoxon.freq[i]
            wilcoxon.cumfreq.append(wilcoxon.var)
        for i in range(len(wilcoxon.cumfreq)):
            for j in range(wilcoxon.freq[i]):
                wilcoxon.total = wilcoxon.total + j
            mean = wilcoxon.total / wilcoxon.freq[i]
            for k in range(wilcoxon.freq[i]):
                wilcoxon.wil_ranks.append(mean)
            wilcoxon.total = wilcoxon.cumfreq[i]
        for i in range(len(wilcoxon.cl_no_zero_diff)):
            if wilcoxon.cl_no_zero_diff[i][0] == 1:
                wilcoxon.pos_sum = wilcoxon.pos_sum + wilcoxon.wil_ranks[i]
            else:
                wilcoxon.neg_sum = wilcoxon.neg_sum + wilcoxon.wil_ranks[i]
        if wilcoxon.pos_sum > wilcoxon.neg_sum:
            print("Wilcoxon signed rank test statistic for",int(sheet.cell_value(0, self.x)),"and",int(sheet.cell_value(0, self.y)),"is",round(wilcoxon.pos_sum, 2))
        else:
            print("Wilcoxon signed rank test statistic for ",int(sheet.cell_value(0, self.x)),"and",int(sheet.cell_value(0, self.y)),"is",round(wilcoxon.neg_sum, 2))
       
wl = wilcoxon("Algo Test.xlsx", 0, 1, 6)
read = wl.read_excel()
wl.wilcox_rank(read)


        
    