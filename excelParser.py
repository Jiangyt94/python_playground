
import xlrd
import xlwt

STATE_SUCCESS = 0
N_TARGET_COL = 21
KESHI_ROW = 2
KESHI_COL = 3
NAME_COL = 1

keshi = {}
salary_type = [
    '岗位工资','薪级工资','保留工资', '津补贴','护龄津贴','护龄贴', '独子费',
    '卫生津贴', '收入合计', '标准绩效', '养老', '医保', '住房', '失业保险',
    '个税', '会费', '支出合计', '实领合计'
]
salary_index = {}

class entry:
    def __init__(self):
        self.data = {}
        self.data['name'] = ''
        self.data['keshi'] = ''
        for t in salary_type:
            self.data[t] = 0

    def parseRow(self, rv, keshi_col):
        self.data['name'] = rv[NAME_COL]
        self.data['keshi'] = rv[keshi_col]
        for id,t in enumerate(salary_type):
            self.data[t] = rv[keshi_col + id+1]


class entry_group:
    def __init__(self):
        self.entrys = []

def convert(input, output):
    print('from:'+input+'\n'+'to:'+output)

    data = xlrd.open_workbook(input)
    tables = data.sheets()
    scan_keshi(tables)
    scan_data(tables)
    do_sum()
    output_data(output)

    return STATE_SUCCESS

def scan_keshi(tables):
    keshi_col  = 0
    for table in tables:
        nrows = table.nrows
        ncols = table.ncols
        if ncols < N_TARGET_COL:
            continue
        for idx, val in enumerate(table.row_values(KESHI_ROW)):
            if val == '科室':
                keshi_col = idx
        for row_id in range( KESHI_ROW+1, nrows-1):

            name = table.row_values(row_id)[keshi_col]
            if type(name) == float:
                name = int(name)
            name = str(name).strip(' ')

            if name not in keshi:
                keshi[name] = []
                #print(name + ':' + table.name)

def scan_data(tables):
    keshi_col = 0
    for table in tables:
        nrows = table.nrows
        ncols = table.ncols
        if ncols < N_TARGET_COL:
            continue

        # get keshi index
        for idx, val in enumerate(table.row_values(KESHI_ROW)):
            if val == '科室':
                keshi_col = idx
        assert table.row_values(KESHI_ROW)[1] == '姓名'

        # get data
        for row_id in range(KESHI_ROW + 1, nrows-1):
            rv = table.row_values(row_id)
            keshi_name = rv[keshi_col]
            if type(keshi_name) == float:
                keshi_name = int(keshi_name)
            keshi_name = str(keshi_name).strip(' ')
            assert keshi_name in keshi

            e = entry()
            e.parseRow(rv, keshi_col)
            keshi[keshi_name].append(e)

def do_sum():
    for k in keshi:
        e = entry()
        e.data['name'] = '合计'
        e.data['keshi'] = k
        for person in keshi[k]:
            for t in salary_type:
                val = person.data[t]
                if val=='':
                    val = 0
                #print(t+':'+str(val))
                e.data[t] += float(val)
        keshi[k].append(e)

def output_data(output):
    data = xlwt.Workbook()
    table = data.add_sheet('整理')
    table.write(0,0, '姓名')
    table.write(0,1, '科室')
    for id, t in enumerate(salary_type):
        table.write(0,2+id, t)
    count = 1

    for k in keshi:
        for e in keshi[k]:
            table.write(count, 0, e.data['name'])
            table.write(count, 1, e.data['keshi'])
            for id, s in enumerate(salary_type):
                table.write(count, id+2, e.data[s])
            count += 1
        count += 1
    data.save(output)



if __name__ == '__main__':
    i = '/Users/nekobus/Downloads/test.xls'
    o = '/Users/nekobus/Downloads/test_.xls'

    convert(i,o)