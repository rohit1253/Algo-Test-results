def combinations(iterable, r):
    # combinations('ABCD', 2) --> AB AC AD BC BD CD
    # combinations(range(4), 3) --> 012 013 023 123
    pool = tuple(iterable)
    n = len(pool)
    if r > n:
        return
    indices = list(range(r))
    yield tuple(pool[i] for i in indices)
    while True:
        for i in reversed(range(r)):
            if indices[i] != i + n - r:
                break
        else:
            return
        indices[i] += 1
        for j in range(i+1, r):
            indices[j] = indices[j-1] + 1
        yield tuple(pool[i] for i in indices)
#        a = tuple(pool[i] for i in indices)
 #   return a
def read_excel(filename):
    '''Reads the first sheet in an excel workbook'''
    import xlrd
    wb = xlrd.open_workbook(filename)
    global sheet
    sheet = wb.sheet_by_index(1)
    return sheet

filename = "Algo Test.xlsx"
read_excel(filename)


a = []
n = 5 # Number of non-key columns
for i in range(2, n+1):
    a.append(list(combinations(range(n), i)))
p = []
q = []
y = []
b = 0
for i in range(len(a)):
    b = len(a[i][0])
    for j in range(b-1):
        for k in range(2, sheet.nrows):
            p.append(sheet.cell_value(k, a[i][i][j]))
        break
    for j in range(1, b):
        print(j)
        for k in range(2, sheet.nrows):
            q.append(sheet.cell_value(k, a[i][i][j]))
        break
    break
for i in range(len(p)):
    y.append((p[i]+q[i])/2)

# Here I got the first combination of the non-key elements.
# In the list a I have got various possible combinations of the non key elements
# But due to time constraint was unable to apply wilcoxon signed rank test to all the combinations




