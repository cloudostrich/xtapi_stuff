import pythoncom
from time import sleep
from win32com.client import Dispatch, DispatchWithEvents, getevents
from win32com.client.gencache import EnsureDispatch, EnsureModule
import osbrain
import pandas as pd

Gate = None
ns = None
NotifyTF = None
NotifySprd = None
NotifyTcm = None

class InstrNotify():#getevents('XTAPI.TTInstrNotify')):
    def __init__(self):
        self.agent_notify = None
        self.addr = None
        self.addr_alias = None

    def gen_agent(self, agentname, channelname):
        self.agent_notify = osbrain.run_agent(agentname)
        self.addr = self.agent_notify.bind('PUSH', alias=channelname)
        self.addr_alias = channelname

    def Subscribe(self, pInstr):
        self.AttachInstrument(pInstr)
        pInstr.Open(0)
        print(f'subscribed: {pInstr.Contract}')

    def OnNotifyFound(self, pNotify=None, pInstr=None):
        pInstr = Dispatch(pInstr)        
        print('Found instrument:')
        print('--> Contract: %s' % pInstr.Get('Contract'))
        print('--> Exchange: %s' % pInstr.Get('Exchange'))

    def OnNotifyNotFound(self, pNotify=None, pInstr=None):
        pInstr = Dispatch(pInstr)        
        print('Unable to find instrument')

    def OnNotifyUpdate(self, pNotify=None, pInstr=None):
        pInstr = Dispatch(pInstr)
        contract = pInstr.Get('Contract')
        bid = pInstr.Get('Bid&')
        ask = pInstr.Get('Ask&')
        last = pInstr.Get('Last&')
        lastqty = pInstr.Get('LastQty')
        bidqty = pInstr.Get('BidQty')
        askqty = pInstr.Get('AskQty')
        upd = {'contract':contract, 'bidqty': bidqty, 'bid': bid, 'ask': ask,
               'askqty': askqty, 'last': last, 'lastqty': lastqty}
        self.agent_notify.send(self.addr_alias, upd)

def Connect():
    global NotifyTF, NotifySprd, NotifyTcm, Gate
    #the below is required in order to establish the com-object links
    #that way you don't need to run makepy first
    EnsureModule('{98B8AE14-466F-11D6-A27B-00B0D0F3CCA6}', 0, 1, 0)

    Gate = EnsureDispatch('XTAPI.TTGate')
    NotifyTF = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    NotifyTF.gen_agent('agentTF','channelTF')
##    print('Connected...')
##    NotifySprd = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
##    NotifySprd.gen_agent('agentSPRD','channelSPRD')
##    print(' Connected Spreads...')
    NotifyTcm = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    NotifyTcm.gen_agent('agentTcm','channelTcm')

def log_message(agent, message):
    agent.log_info(f'XTapi: {message}')
    
def qq(ns,Gate):
    Gate.XTAPITerminate()
    ns.shutdown()

def sub_fut(pinstr, exch, prod, contract, notifier):
    pinstr.Exchange = exch
    pinstr.Product = prod
    pinstr.ProdType = 'FUTURE'
    pinstr.Contract = contract
    notifier.Subscribe(pinstr)

def sub_sprd(pinstr, exch, prod, leg1, leg2, notifier):
    pinstr.Exchange = exch
    pinstr.Product = prod
    pinstr.ProdType = 'SPREAD'
    pinstr.Contract = f'Calendar: 1xTF {leg1}:-1xTF {leg2}'
    notifier.Subscribe(pinstr)

def main():
    global ns
    ns = osbrain.run_nameserver()
    tcm = osbrain.run_agent('Tcm')
    bob = osbrain.run_agent('Bob')
##    sprd = osbrain.run_agent('Sprd')
    Connect()
    tcm.connect(NotifyTcm.addr, handler=log_message)
    bob.connect(NotifyTF.addr, handler=log_message)
##    sprd.connect(NotifySprd.addr, handler=log_message)

    ''' init tocom rss
    =RTD("XTAPI.RTD","","Instr","TOCOM-B","RSS3","FUTURE","25Jan19","ASK")
    =RTD("XTAPI.RTD","","Instr","SGX-B","TF","FUTURE","Feb19","ASK")
    '''
    tcm_25Jan19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_22Feb19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_25Mar19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_22Apr19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_27May19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_24Jun19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_mth = ['25Jan19', '22Feb19', '25Mar19',
               '22Apr19', '27May19', '24Jun19']
    tcm_lst = [tcm_25Jan19, tcm_22Feb19, tcm_25Mar19,
               tcm_22Apr19, tcm_27May19, tcm_24Jun19]
    
    # dispatch tocom rss
    for tl, tm in zip(tcm_lst, tcm_mth):
        sub_fut(tl, 'TOCOM-B', 'RSS3', tm,NotifyTcm)
    NotifyTcm.UpdateFilter = 'Bid, Ask, Last, LastQty'
##    sub_fut(tcm_25Jan19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm)
##    sub_fut(tcm_22Feb19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm)
##    sub_fut(tcm_25Mar19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm)
##    sub_fut(tcm_22Apr19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm)
##    sub_fut(tcm_27May19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm)
##    sub_fut(tcm_24Jun19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm)

    # dispatch tf outrights
##    mar19 = EnsureDispatch('XTAPI.TTInstrObj')
##    apr19 = EnsureDispatch('XTAPI.TTInstrObj')
##    may19 = EnsureDispatch('XTAPI.TTInstrObj')
##    jun19 = EnsureDispatch('XTAPI.TTInstrObj')
##    jul19 = EnsureDispatch('XTAPI.TTInstrObj')
##    aug19 = EnsureDispatch('XTAPI.TTInstrObj')
##    sep19 = EnsureDispatch('XTAPI.TTInstrObj')
##    oct19 = EnsureDispatch('XTAPI.TTInstrObj')
##    nov19 = EnsureDispatch('XTAPI.TTInstrObj')
##    dec19 = EnsureDispatch('XTAPI.TTInstrObj')
##    jan20 = EnsureDispatch('XTAPI.TTInstrObj')
##    
##    sub_fut(mar19, 'SGX-B', 'TF', 'Mar19',NotifyTF)
##    sub_fut(apr19, 'SGX-B', 'TF', 'Apr19',NotifyTF)
##    sub_fut(may19, 'SGX-B', 'TF', 'May19',NotifyTF)
##    sub_fut(jun19, 'SGX-B', 'TF', 'Jun19',NotifyTF)
##    sub_fut(jul19, 'SGX-B', 'TF', 'Jul19',NotifyTF)
##    sub_fut(aug19, 'SGX-B', 'TF', 'Aug19',NotifyTF)
##    sub_fut(sep19, 'SGX-B', 'TF', 'Sep19',NotifyTF)
##    sub_fut(oct19, 'SGX-B', 'TF', 'Oct19',NotifyTF)
##    sub_fut(nov19, 'SGX-B', 'TF', 'Nov19',NotifyTF)
##    sub_fut(dec19, 'SGX-B', 'TF', 'Dec19',NotifyTF)
##    sub_fut(jan20, 'SGX-B', 'TF', 'Jan20',NotifyTF)
##    
##    NotifyTF.UpdateFilter = 'Bid, Ask, Last, LastQty'
##
##    # Dispatch TF spreads
##    h2 = EnsureDispatch('XTAPI.TTInstrObj')
##    h3 = EnsureDispatch('XTAPI.TTInstrObj')
##    h4 = EnsureDispatch('XTAPI.TTInstrObj')
##    h5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    j2 = EnsureDispatch('XTAPI.TTInstrObj')
##    j3 = EnsureDispatch('XTAPI.TTInstrObj')
##    j4 = EnsureDispatch('XTAPI.TTInstrObj')
##    j5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    k2 = EnsureDispatch('XTAPI.TTInstrObj')
##    k3 = EnsureDispatch('XTAPI.TTInstrObj')
##    k4 = EnsureDispatch('XTAPI.TTInstrObj')
##    k5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    m2 = EnsureDispatch('XTAPI.TTInstrObj')
##    m3 = EnsureDispatch('XTAPI.TTInstrObj')
##    m4 = EnsureDispatch('XTAPI.TTInstrObj')
##    m5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    n2 = EnsureDispatch('XTAPI.TTInstrObj')
##    n3 = EnsureDispatch('XTAPI.TTInstrObj')
##    n4 = EnsureDispatch('XTAPI.TTInstrObj')
##    n5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    q2 = EnsureDispatch('XTAPI.TTInstrObj')
##    q3 = EnsureDispatch('XTAPI.TTInstrObj')
##    q4 = EnsureDispatch('XTAPI.TTInstrObj')
####    q5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    u2 = EnsureDispatch('XTAPI.TTInstrObj')
##    u3 = EnsureDispatch('XTAPI.TTInstrObj')
####    u4 = EnsureDispatch('XTAPI.TTInstrObj')
####    u5 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    v2 = EnsureDispatch('XTAPI.TTInstrObj')
##
##    x2 = EnsureDispatch('XTAPI.TTInstrObj')

    # subsribe spreads
    #=RTD("XTAPI.RTD","","Instr","SGX-B","TF","SPREAD","Calendar: 1xTF Aug19:-1xTF Nov19","ASK")
##    sub_sprd(h2, 'SGX-B', 'TF', 'Mar19', 'May19', NotifySprd)
##    sub_sprd(h3, 'SGX-B', 'TF', 'Mar19', 'Jun19', NotifySprd)
##    sub_sprd(h4, 'SGX-B', 'TF', 'Mar19', 'Jul19', NotifySprd)
##    sub_sprd(h5, 'SGX-B', 'TF', 'Mar19', 'Aug19', NotifySprd)
##
##    sub_sprd(j2, 'SGX-B', 'TF', 'Apr19', 'Jun19', NotifySprd)
##    sub_sprd(j3, 'SGX-B', 'TF', 'Apr19', 'Jul19', NotifySprd)
##    sub_sprd(j4, 'SGX-B', 'TF', 'Apr19', 'Aug19', NotifySprd)
##    sub_sprd(j5, 'SGX-B', 'TF', 'Apr19', 'Sep19', NotifySprd)
##
##    sub_sprd(k2, 'SGX-B', 'TF', 'May19', 'Jul19', NotifySprd)
##    sub_sprd(k3, 'SGX-B', 'TF', 'May19', 'Aug19', NotifySprd)
##    sub_sprd(k4, 'SGX-B', 'TF', 'May19', 'Sep19', NotifySprd)
##    sub_sprd(k5, 'SGX-B', 'TF', 'May19', 'Oct19', NotifySprd)
##
##    sub_sprd(m2, 'SGX-B', 'TF', 'Jun19', 'Aug19', NotifySprd)
##    sub_sprd(m3, 'SGX-B', 'TF', 'Jun19', 'Sep19', NotifySprd)
##    sub_sprd(m4, 'SGX-B', 'TF', 'Jun19', 'Oct19', NotifySprd)
##    sub_sprd(m5, 'SGX-B', 'TF', 'Jun19', 'Nov19', NotifySprd)
##
##    sub_sprd(n2, 'SGX-B', 'TF', 'Jul19', 'Sep19', NotifySprd)
##    sub_sprd(n3, 'SGX-B', 'TF', 'Jul19', 'Oct19', NotifySprd)
##    sub_sprd(n4, 'SGX-B', 'TF', 'Jul19', 'Nov19', NotifySprd)
##    sub_sprd(n5, 'SGX-B', 'TF', 'Jul19', 'Dec19', NotifySprd)
##
##    sub_sprd(q2, 'SGX-B', 'TF', 'Aug19', 'Oct19', NotifySprd)
##    sub_sprd(q3, 'SGX-B', 'TF', 'Aug19', 'Nov19', NotifySprd)
##    sub_sprd(q4, 'SGX-B', 'TF', 'Aug19', 'Dec19', NotifySprd)
##
##    sub_sprd(u2, 'SGX-B', 'TF', 'Sep19', 'Nov19', NotifySprd)
##    sub_sprd(u3, 'SGX-B', 'TF', 'Sep19', 'Dec19', NotifySprd)
##
##    sub_sprd(v2, 'SGX-B', 'TF', 'Oct19', 'Dec19', NotifySprd)
##
##    sub_sprd(x2, 'SGX-B', 'TF', 'Nov19', 'Jan20', NotifySprd)
##
##    NotifyTF.UpdateFilter = 'Bid, Ask, Last, LastQty'

    for i in range(15):
        print('pumping...')
        pythoncom.PumpWaitingMessages()
        sleep(1.0)
