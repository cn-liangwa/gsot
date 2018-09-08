import xml.etree.ElementTree as ET
import cx_Oracle
import re
import xlwt
import html
import configparser


def main():
##    tree = ET.parse('package.xml')
##    root = tree.getroot()
##    print(root)
    openDBSession()

def openDBSession():
    cf = configparser.ConfigParser()
    cf.read('../gsot.cfg')
    
    db_host = cf.get('db','host')
    db_port = cf.get('db','port')
    db_service_name = cf.get('db','service_name')
    db_user = cf.get('db','user')
    db_password = cf.get('db','password')
      
    procdef_map = None
    
    dsn = cx_Oracle.makedsn(db_host,db_port,db_service_name)
    conn = cx_Oracle.connect(db_user,db_password,dsn)   #注意，这是生产环境数据库!!!!!!
    curs=conn.cursor()
    
    procdef_map = getProcdefMap(curs)  
    
    
    book = xlwt.Workbook()
    sheet1 = book.add_sheet(u'sheet1',cell_overwrite_ok=True)
    
    sheet1.write(0,0,'单位编码')
    sheet1.write(0,1,'单位名称')
    sheet1.write(0,2,'流程名称')
    sheet1.write(0,3,'流程配置')
    
         
    qry_str = """
                    SELECT lsbzdw.lsbzdw_dwbh          AS dwbh,
                       lsbzdw.lsbzdw_dwmc          AS dwmc,
                       package.name                AS package_name,
                       procdef.processdefinitionid AS procdefid
                  FROM lsbzdw lsbzdw
                 INNER JOIN processassigns rel_proc_comp
                    ON lsbzdw.lsbzdw_dwbh = rel_proc_comp.companyid
                 INNER JOIN bizprocess proc
                    ON rel_proc_comp.processid = proc.id
                 INNER JOIN relationbizprocprocdef rel_proc_procdef
                    ON rel_proc_comp.processid = rel_proc_procdef.bizprocid
                 INNER JOIN processdefinition procdef
                    ON rel_proc_procdef.procdefid = procdef.processdefinitionid
                 INNER JOIN package package
                    ON procdef.packageid = package.packageid
                 WHERE procdef.state = '1'
                   AND lsbzdw.lsbzdw_mx = '1'
                   AND lsbzdw.lsbzdw_dwbh in('5301')
                  
    """        
    curs.execute(qry_str)
    rss = curs.fetchall()

    for rownum,row in enumerate(rss,start=1):
        sheet1.write(rownum,0,row[0])
        sheet1.write(rownum,1,row[1])
        sheet1.write(rownum,2,row[2])
        procdef = procdef_map.get(row[3])
        ps = procdef.findPath(row[0])
        for i,p in enumerate(ps,start=3):
           sheet1.write(rownum,i,p) 
           
    book.save('out.xls')
    
    curs.close();
    conn.close();
    
            
def getProcdefMap(curs):
    
    procdef_dict = {}
    procdef = None
    rss_procdef = None
    sqlstr="""
    
            SELECT procdef.processdefinitionid, procdef.packageid, package.xpdlpackage
              FROM processdefinition procdef
             INNER JOIN PACKAGE
                ON procdef.packageid = package.packageid
             WHERE procdef.state = '1'
             
    """    
    curs.execute(sqlstr)
    rss_procdef =curs.fetchall()
    
    for row in rss_procdef:       
        xmlstr =  getXmlstrByXPDL(row[2])
        procdef = getProcDef(row[0],xmlstr)
        procdef_dict[row[0]] = procdef
        
    return procdef_dict

def getXmlstrByXPDL(XPDL):
    xmlstr = None
    if isinstance(XPDL,cx_Oracle.LOB):
        chunksize = XPDL.getchunksize()
        clobsize = XPDL.size()
        nchunk = clobsize-clobsize%chunksize
        xmlstr = str(XPDL.read(amount=nchunk),'gb2312',errors = 'ignore')
        xmlstr +=  str(XPDL.read(offset=nchunk+1),'gb2312',errors = 'ignore')
        xmlstr = xmlstr.replace('gb2312', 'utf-8')
    return xmlstr


def getProcDef(procdefid,xmlstr):
    
    ns = {'xpdl':'http://www.wfmc.org/2002/XPDL1.0',
          'xsi':'http://www.w3.org/2001/XMLSchema-instance',
          'gsp':'http://www.genersoft.com/GSP1.0'}
    
    root = ET.fromstring(xmlstr)
    workflowprocess_root = root.find(".//xpdl:WorkflowProcess[@Id='{0}']".format(procdefid),namespaces=ns)
    participants = {}
    activities   = {}
    transitions  = {}


    #提取流程图所有参与者信息
    for p in workflowprocess_root.iter('{http://www.wfmc.org/2002/XPDL1.0}Participant'):
        participants[p.get('Id')]= {
            'code':p.get('Code'),
            'name':p.get('Name')
            }
        
    #提取审批节点信息
    for a in workflowprocess_root.iter('{http://www.wfmc.org/2002/XPDL1.0}Activity'):

        #首先提取节点审批人员列表
        performers = {}
        performerNodeText = ''
        performerNode = a.find('{http://www.wfmc.org/2002/XPDL1.0}Performer')
        if performerNode is not None:
            performerNodeText = performerNode.text
        if performerNodeText is not None and performerNodeText != '':
            for pid in performerNodeText.split(','):
                performers[pid] = {'condition':''}
        if len(performers)>0:                       
            for pc in a.iter('{http://www.genersoft.com/GSP1.0}Performer'):
                condition = pc.find('{http://www.genersoft.com/GSP1.0}Condition')
                performer = performers.get(pc.get('Id'))
                if performer is not None and condition is not None:
                    performer['condition']=  condition.get('Value',default='')
        
        implementationType = ''
        implementationToolNode = a.find('./xpdl:Implementation/xpdl:Tool',ns)
        if implementationToolNode:
            ID = implementationToolNode.get('Id')
            if ('Refuse' in ID) or ('NoPass' in ID) or ('NotPass' in ID):           
                implementationType = 'refuse'
            else: 
                implementationType = 'pass'
             
        transitionRefs = []
        for tr in a.iter('{http://www.wfmc.org/2002/XPDL1.0}TransitionRef'):
            transitionRefs.append(tr.get('Id'))
            
        
        activityType = ''
        if '[StartActivity]' in a.get('Id'):
            activityType = 'start'
        elif '[EndActivity]' in a.get('Id'):
            activityType = 'end'
        elif '[AutoActivity]' in a.get('Id'):
            activityType = 'auto'
        elif '[RouteActivity]' in a.get('Id'):
            activityType = 'route'
        elif '[ManualActivity]' in a.get('Id'):
            activityType = 'manual'
        else:
            activityType = 'unknown'
            
        activities[a.get('Id')] = {
            'name':a.get('Name'),
            'type':activityType,
            'performers':performers,
            'transitionRefs':transitionRefs,
            'implementationType':implementationType
            }

    #提取所有箭头信息
    for t in workflowprocess_root.iter('{http://www.wfmc.org/2002/XPDL1.0}Transition'):
        condition = t.find('{http://www.wfmc.org/2002/XPDL1.0}Condition')
        transitions[t.get('Id')] = {
            'to':t.get('To'),
            'from':t.get('From'),
            'conditionType':condition.get('Type'),
            'conditionValue':condition.get('Value')
            }
        
    return processDef(participants,activities,transitions)


class processDef:

    def __init__(self,participants,activities,transitions):
        self.participants = participants
        self.activities = activities
        self.transitions = transitions

    def findPath(self,unit):
        paths = []
        path = ''
        for k in self.transitions:
            if self.transitions[k].get('from') == '[StartActivity]startActivity':         
                paths = self.recurfind(path,self.transitions[k],unit)      
        return paths
    
    
    def recurfind(self,path,transition,unit):
        
        nextActivityName = transition.get('to')        
        paths = []
        atv = self.activities[nextActivityName]
        
        #如果结束审批或者审批通过，则视为完整路径
        if atv.get('type') == 'end' or atv.get('type') == 'auto':                
            return paths
        
        #在路径中加上分支转移的特殊条件
        t_condition_value = self.get_transitionRef_condition(transition.get('conditionValue')) 
        path += '--' + t_condition_value + '--'

        #在路径中加上节点名称
        path += '[' + atv.get('name') + ']'
        
        #如果为审批节点，则加上审批人员信息
        if atv.get('type') == 'manual':               
                performers = atv.get('performers')
                if len(performers)>0:                   
                    numPerformers = 0
                    for k in performers:                   
                        if self.calculateCondition_a(unit,performers[k].get('condition')):
                            p = self.participants.get(k)
                            path += ('(' + p.get('code') + ',' + p.get('name') + ')')
                            numPerformers +=1
                    if numPerformers==0:
                        path+='*没有配置审批人员*'
                        
                        
        #获取[满足条件]分支线，[不满足则转移]分支线
        valid_transitionRefs = []
        
        for t in atv.get('transitionRefs'):                 
            if self.calculateCondition_t(unit,self.transitions[t].get('conditionValue')):
                    valid_transitionRefs.append(self.transitions[t])           
 
        if len(valid_transitionRefs) == 0:
            path += '*没有合适的分支线*'
            
        for t in valid_transitionRefs:
            paths.extend(self.recurfind(path,t,unit))
            
        if len(paths)==0:
            paths.append(path)
       
        return paths

    def get_transitionRef_condition(self,condition_value):
        expression_str = condition_value.strip()
        expression_str = expression_str[4:-5]
        res = ''
        if expression_str != '':            
            expression_xml = html.unescape(html.unescape(expression_str))
            m = re.search('金额\s*[<>!=]{1,2}\s*[0-9\.]+',expression_xml)
            if m:
                res = m.group(0)
        return res

            
        


    def calculateCondition_t(self,unit,conditionStr):
        spjl0 = conditionStr.find('审批结论==0') != -1
        spjl1 = conditionStr.find('审批结论==1') != -1
        dwbh  = conditionStr.find('单位')    != -1
        
        if (dwbh and re.search('[^0-9a-zA-Z]'+ unit + '[^0-9a-zA-Z]',conditionStr)) or (not dwbh and not spjl0):
            return True
        else:
            return False


    def calculateCondition_a(self,unit,conditionStr):
        """
        
        
        """
        if re.search('[^0-9a-zA-Z]'+ unit + '[^0-9a-zA-Z]',conditionStr) or not conditionStr.find('单位')!= -1:
            return True
        else:
            return False

        
        

        
    
    


##def downLoadWorkFlow():
##    
##
##
##def upLoadWorkFlow():
##
##    

if __name__=='__main__' :
    main()
