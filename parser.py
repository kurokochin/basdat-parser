import datetime
import sys

def minimalist_xldate_as_datetime(xldate, datemode):
    # datemode: 0 for 1900-based, 1 for 1904-based
    return (
        datetime.datetime(1899, 12, 30)
        + datetime.timedelta(days=int(xldate) + 1462 * datemode)
        )

def xlsx(fname):
    import zipfile
    from xml.etree.ElementTree import iterparse
    z = zipfile.ZipFile(fname)
    strings = [el.text for e, el in iterparse(z.open('xl/sharedStrings.xml')) if el.tag.endswith('}t')]
    rows = []
    row = {}
    value = ''
    ws = 'xl/worksheets/' + sys.argv[2] + '.xml'
    for e, el in iterparse(z.open(ws)):
        if el.tag.endswith('}v'): # <v>84</v>
            value = el.text
        if el.tag.endswith('}c'): # <c r="A3" t="s"><v>84</v></c>
            if el.attrib.get('t') == 's':
                value = strings[int(value)]
            letter = el.attrib['r'] # AZ22
            while letter[-1].isdigit():
                letter = letter[:-1]
            row[letter] = value
        if el.tag.endswith('}row'):
            rows.append(row)
            row = {}
    return rows

def query(result):
    key_date = []
    sql_query = "INSERT INTO " + sys.argv[4] + " ("
    first = True
    index = 0
    for i in sorted(result[0]):
        if result[0][i].replace(' ', '') == '':
            continue
        if result[0][i].find("tgl") != -1 or result[0][i].find("tanggal") != -1:
            key_date.append(index)
        if first:
            sql_query += result[0][i]
            first = False
        else:
            sql_query += ', ' + result[0][i]
        index += 1
    sql_query += ') VALUES ( %s );'
    queries = []
    for i in result[1:]:
        value = ''
        first = True
        index = 0
        for j in sorted(i):
            if first:
                if index in key_date and type(i[j]) == type(1):
                    value += '\'' + str(minimalist_xldate_as_datetime(i[j], 0).date()) + '\''
                else:
                    if i[j] != 'NULL':
                        value += '\'' + i[j] + '\''
                    else:
                        value += 'NULL'
                first = False
            else:
                if index in key_date and type(i[j]) == type(1):
                    value += ' ,\'' + str(minimalist_xldate_as_datetime(i[j], 0).date()) + '\''
                else:
                    if i[j] != 'NULL':
                        value += ' ,\'' + i[j] + '\''
                    else:
                        value += ', NULL'
            index += 1
        queries.append(sql_query % value)
    f = open(sys.argv[3], 'w')
    for i in queries:
        f.write(i+'\n')
    f.close()

query(xlsx(sys.argv[1]))
