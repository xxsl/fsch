VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'    private void readCounts() throws IOException
'    {
'        int classCount = read32();
'
'        for (int i = 0; i < classCount; i++)
'        {
'            String type = this.playerSession.getString(read32());
'            int functionCount = read32();
'
'            for (int j = 0; j < functionCount; j++)
'            {
'                String functionName = this.playerSession.getString(read32());
'                long invocationCount = read64();
'                this.playerSession.counts(QNameUtil.typeToQName(type), QNameUtil.memberToQName(functionName), invocationCount);
'            }
'
'        }
'
'        this.playerSession.countsDone();
'    }
