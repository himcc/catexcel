# catexcel

> print excel file to console.

## Usage
```sh
# catexcel -h
cat excel file
catexcel a.xlsx b.xlsx               # show top 10 lines for every sheet
catexcel -s 0 -s 1 a.xlsx            # show complete content in first sheet and second sheet
catexcel -f csv -s students a.xlsx   # show complete content in students sheet with csv format(default tsv)

Options:

  -h, --help           display help information
  -a, --all[=false]    output all sheets
  -f, --format[=tsv]   tsv or csv
  -s, --sheet          sheet name or index (from 0) for output
```

## Example
```sh
# catexcel students.xlsx teacher.xlsx
>> students.xlsx
>> 0 : Sheet1
No      Name
1       odin
2       amy
3       cora
4       suzanne
5       jini
6       katherine
7       candice
8       selena
9       abigail

>> students.xlsx
>> 1 : Sheet2
a       b       c       d
e       f       g       i
j       k       l       m

>> students.xlsx
>> 2 : Sheet3

>> teacher.xlsx
>> 0 : Sheet1
No      Name
1       zhangsan
2       lisi
3       wangwu
4       zhaoliu
5       sunqi

>> teacher.xlsx
>> 1 : Sheet2

>> teacher.xlsx
>> 2 : Sheet3
```

```sh
# catexcel -s 0 -s Sheet2 -f csv students.xlsx
No,Name
1,odin
2,amy
3,cora
4,suzanne
5,jini
6,katherine
7,candice
8,selena
9,abigail
10,elsa
11,vicky
12,hanna
13,zora
14,grace
15,madge
16,felicia
17,gina
18,paula
19,mervin
20,verna
a,b,c,d
e,f,g,i
j,k,l,m
```

```sh
# catexcel -a students.xlsx
No      Name
1       odin
2       amy
3       cora
4       suzanne
5       jini
6       katherine
7       candice
8       selena
9       abigail
10      elsa
11      vicky
12      hanna
13      zora
14      grace
15      madge
16      felicia
17      gina
18      paula
19      mervin
20      verna
a       b       c       d
e       f       g       i
j       k       l       m
```

## Cross compile (in Windows PowerShell)
```sh
$Env:GOOS = "darwin"; $Env:GOARCH = "amd64"; go build -o catexcel-darwin-amd64
$Env:GOOS = "linux"; $Env:GOARCH = "amd64"; go build -o catexcel-linux-amd64
```