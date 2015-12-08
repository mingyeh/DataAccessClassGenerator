# -*- coding:UTF-8 -*-
import string,os,sys,xlrd

dbUser = 'REFORM'
tableSpace = 'REFORM_DATA'
docsDir = r'd:\demo\tables\docs' # Directory of database documents
sqlDir = r'd:\demo\tables\sqls'
entityDir = r'd:\demo\tables\entities'

default_encoding = 'utf-8'
if sys.getdefaultencoding() != default_encoding:
    reload(sys)
    sys.setdefaultencoding(default_encoding)

def columnNameCast(columnName):
	rArr = []
	cArr = columnName.split('_')
	for s in cArr:
		rArr.append(s.title())
	return ''.join(rArr)

def columnNameCamal(columnName):
	rArr = []
	cArr = columnName.split('_')
	for s in cArr:
		if cArr.index(s) == 0:
			rArr.append(s.lower())
		else:
			rArr.append(s.title())
	return ''.join(rArr)

stringCode = '''
	/**
	 * $desc$
	 */
	@Column(name="$columnName$")
	@Length(min=0,max=$len$,message="$desc$最多$len$个汉字")
	private String $castName$;
	
	/**
	 * @return $desc$
	 */
	public String get$castName$() {
		return $castName$;
	}

	/**
	 * @param 设置的$desc$
	 */
	public void set$castName$(String $camalName$) {
		$castName$ = $camalName$;
	}
'''

dateCode = '''
	/**
	 * $desc$
	 */
	@Column(name = "$columnName$")
	@Past(message="日期必须小于当前日期")
	private Date $castName$;
	
	/**
	 * @return $desc$
	 */
	public Date get$castName$() {
		return $castName$;
	}

	/**
	 * @param 设置的$desc$
	 */
	public void set$castName$(Date $camalName$) {
		$castName$ = $camalName$;
	}
'''
numberCode = '''
	/**
	 * $desc$
	 */
	@Column(name = "$columnName$")
	private $numberType$ $castName$;
	
	/**
	 * @return $desc$
	 */
	public $numberType$ get$castName$() {
		return $castName$;
	}

	/**
	 * @param 设置的$desc$
	 */
	public void set$castName$($numberType$ $camalName$) {
		$castName$ = $camalName$;
	}
'''

class columnInfo(object):
	columnName = '' # column name
	chineseName = '' # table description
	dataType = '' # data type
	columnLength = '' # column length
	nullable = True # nullable
	isKey = False # is primary key

class tableInfo(object):
	tableName = '' # table name
	chineseName = '' # table description
	
	def __init__(self):
		self.columns = [] # column list

	def addColumn(self,column):
		self.columns.append(column)
	
	def getColumns(self):
		return self.columns

	def writeSQL(self):
		# write table create sql script
		sqlFile = open(sqlDir + '\\' + self.tableName + '.sql','w')
		#write create sql script	
		sqlFile.write('''
CREATE TABLE "%s"."%s" 
   (''' % (dbUser,self.tableName))

		for col in self.columns:
			#print col.columnName + ' --->> ' + col.chineseName
			if col.dataType.upper() == 'VARCHAR2':
				sqlFile.write('\t"%s" %s(%s BYTE)' % (col.columnName,col.dataType.upper(),col.columnLength))
			elif col.dataType.upper() == 'DATE':
				sqlFile.write('\t"%s" %s' % (col.columnName,col.dataType.upper()))
			elif col.dataType.upper() == 'NUMBER':
				if col.columnLength == '':
					sqlFile.write('\t"%s" %s' % (col.columnName,col.dataType.upper()))
				else:
					sqlFile.write('\t"%s" %s(%s)' % (col.columnName,col.dataType.upper(),col.columnLength))

			if self.columns.index(col) < len(self.columns) -1:
				sqlFile.write(',\n')

		sqlFile.write('''
) SEGMENT CREATION IMMEDIATE 
  PCTFREE 10 PCTUSED 40 INITRANS 1 MAXTRANS 255 
 NOCOMPRESS LOGGING
  STORAGE(INITIAL 65536 NEXT 1048576 MINEXTENTS 1 MAXEXTENTS 2147483645
  PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1
  BUFFER_POOL DEFAULT FLASH_CACHE DEFAULT CELL_FLASH_CACHE DEFAULT)
  TABLESPACE "%s" ;\n\n''' % tableSpace)
		
		#write comment script
		for col in self.columns:
			sqlFile.write('''COMMENT ON COLUMN "%s"."%s"."%s" IS '%s';\n''' % (dbUser,self.tableName,col.columnName,col.chineseName))

		#create primary key
		sqlFile.write('\n')
		pkList = filter(lambda x:x.isKey == True,self.columns)
		for pk in pkList:
			sqlFile.write('''
  CREATE UNIQUE INDEX "%s"."%s" ON "%s"."%s" ("%s") 
  PCTFREE 10 INITRANS 2 MAXTRANS 255 COMPUTE STATISTICS 
  STORAGE(INITIAL 65536 NEXT 1048576 MINEXTENTS 1 MAXEXTENTS 2147483645
  PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1
  BUFFER_POOL DEFAULT FLASH_CACHE DEFAULT CELL_FLASH_CACHE DEFAULT)
  TABLESPACE "%s" ;\n
''' % (dbUser,pk.columnName,dbUser,self.tableName,pk.columnName,tableSpace))
			sqlFile.write('''
  ALTER TABLE "%s"."%s" ADD CONSTRAINT "%s" PRIMARY KEY ("%s")
  USING INDEX PCTFREE 10 INITRANS 2 MAXTRANS 255 COMPUTE STATISTICS 
  STORAGE(INITIAL 65536 NEXT 1048576 MINEXTENTS 1 MAXEXTENTS 2147483645
  PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1
  BUFFER_POOL DEFAULT FLASH_CACHE DEFAULT CELL_FLASH_CACHE DEFAULT)
  TABLESPACE "%s"  ENABLE;
  ALTER TABLE "%s"."%s" MODIFY ("%s" NOT NULL ENABLE);
''' % (dbUser,self.tableName,pk.columnName,pk.columnName,tableSpace,dbUser,self.tableName,pk.columnName))

		#create not null constraints
		sqlFile.write('\n\n')
		notNullCols = filter(lambda x:x.isKey == False and x.nullable == False,self.columns)
		for notNullCol in notNullCols:
			sqlFile.write('''ALTER TABLE "%s"."%s" MODIFY ("%s" NOT NULL ENABLE);\n''' % (dbUser,self.tableName,notNullCol.columnName))

		sqlFile.close()
	
	def writeEntity():
		className = columnNameCast(self.tableName)
		entityFile = open(entityDir + '\\' + className + '.java')
		entityFile.close()

def readDatabaseDocument(docPath):
	book = xlrd.open_workbook(docPath) # Open document
	totalSheetCount = len(book.sheets())
	# get all data table sheets, the first and second sheet are ignored
	for i in range(2,totalSheetCount):
		sheet = book.sheet_by_index(i) # get data table sheet
		tableChineseName = sheet.cell_value(4,4) # get table description
		tableEnglishName = sheet.cell_value(3,4) # get table name
		if tableChineseName == '' or tableEnglishName == '':
			break;
		totalRowCount = sheet.nrows
		
		# create table info object
		table = tableInfo()
		table.tableName = tableEnglishName
		table.chineseName = tableChineseName

		# loop throwgh columns rows
		for j in range(10,totalRowCount):
			seq = sheet.cell_value(j,2)
			if seq != None and seq != '':
				column = columnInfo()
				column.columnName = sheet.cell_value(j,2)
				column.chineseName = sheet.cell_value(j,7)
				column.dataType = sheet.cell_value(j,12)
				column.columnLength = str(sheet.cell_value(j,16)).replace('(','').replace(')','').replace('（','').replace('）','')
				column.nullable = sheet.cell_value(j,19) == 'YES'
				column.isKey = sheet.cell_value(j,22) == "YES"
				table.addColumn(column)
			else:
				break;
		print 'columns.count:' + str(len(table.columns))
		table.writeSQL()

for root,dirs,files in os.walk(docsDir):
	for name in files:
		if name != '':
			print os.path.join(root,name)
			readDatabaseDocument(os.path.join(root,name))

