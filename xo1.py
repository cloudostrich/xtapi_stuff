import pythoncom
from time import sleep
from win32com.client import Dispatch, DispatchWithEvents, getevents
from win32com.client.gencache import EnsureDispatch, EnsureModule
import osbrain

Gate = None
Notify = None
ns = None

class InstrNotify():#getevents('XTAPI.TTInstrNotify')):
    def __init__(self):
        self.alice = osbrain.run_agent('Alice')
        self.addr = self.alice.bind('PUSH', alias='main')

    def Subscribe(self, pInstr):
        self.AttachInstrument(pInstr)
        pInstr.Open(0)
        print(f'subscribed: {pInstr}')

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

        bid = pInstr.Get('Bid')
        ask = pInstr.Get('Ask')
        last = pInstr.Get('Last')
        lastqty = pInstr.Get('LastQty')
        bidqty = pInstr.Get('BidQty')
        askqty = pInstr.Get('AskQty')
##        self.UpdateFilter = 'Bid', 'Ask', 'Last', 'LastQty'

##        upd = f'[UPDATE] {contract}: {bidqty} | {bid} / {ask} | {askqty} :: {last} | {lastqty}'
        upd = {'contract':contract, 'bidqty': bidqty, 'bid': bid, 'ask': ask, 'askqty': askqty, 'last': last, 'lastqty': lastqty}
        self.alice.send('main', upd)
##        print('[UPDATE] %s: %s/%s' % (contract, bid, ask))


def Connect():
    global Notify, Gate
    #the below is required in order to establish the com-object links
    #that way you don't need to run makepy first
    EnsureModule('{98B8AE14-466F-11D6-A27B-00B0D0F3CCA6}', 0, 1, 0)

    Gate = EnsureDispatch('XTAPI.TTGate')
    Notify = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    print('Connected...')
##    NOTIFY = Dispatch('XTAPI.TTInstrNotify', InstrNotify)
    
def log_message(agent, message):
    agent.log_info(f'From XTapi: {message}')
    
def qq(ns,Gate):
    Gate.XTAPITerminate()
    ns.shutdown()

def sub_tf(pinstr, contract):
    pinstr.Exchange = 'SGX-B'
    pinstr.Product = 'TF'
    pinstr.ProdType = 'FUTURE'
    pinstr.Contract = contract
    Notify.Subscribe(pinstr)
    
def main():
    global ns
    ns = osbrain.run_nameserver()
    bob = osbrain.run_agent('Bob')
    Connect()
    bob.connect(Notify.addr, handler=log_message)
##    pInstr = EnsureDispatch('XTAPI.TTInstrObj')
##    pInstr.Exchange = 'TOCOM-B'
##    pInstr.Product  = 'RSS3'
##    pInstr.Contract = '24Jun19'
##    pInstr.ProdType = 'FUTURE'
##    Notify.Subscribe(pInstr)
    mar19 = EnsureDispatch('XTAPI.TTInstrObj')
    apr19 = EnsureDispatch('XTAPI.TTInstrObj')
    may19 = EnsureDispatch('XTAPI.TTInstrObj')
    jun19 = EnsureDispatch('XTAPI.TTInstrObj')
    jul19 = EnsureDispatch('XTAPI.TTInstrObj')
    aug19 = EnsureDispatch('XTAPI.TTInstrObj')
    sep19 = EnsureDispatch('XTAPI.TTInstrObj')
    oct19 = EnsureDispatch('XTAPI.TTInstrObj')
    nov19 = EnsureDispatch('XTAPI.TTInstrObj')
    dec19 = EnsureDispatch('XTAPI.TTInstrObj')

    sub_tf(mar19, 'Mar19')
    sub_tf(apr19, 'Apr19')
    sub_tf(may19, 'May19')
    sub_tf(jun19, 'Jun19')
    sub_tf(jul19, 'Jul19')
    sub_tf(aug19, 'Aug19')
    sub_tf(sep19, 'Sep19')
    sub_tf(oct19, 'Oct19')
    sub_tf(nov19, 'Nov19')
    sub_tf(dec19, 'Dec19')
    Notify.UpdateFilter = 'Bid, Ask, Last, LastQty'


    for i in range(10):
        print('pumping...')
        pythoncom.PumpWaitingMessages()
        sleep(1.0)

    def qqm():
        Gate.XTAPITerminate()
        ns.shutdown()
        

"""
Found instrument:
--> Contract: CL Mar13
--> Exchange: CME-A
[UPDATE] CL Mar13: 9760/9764
=RTD("XTAPI.RTD","","Instr","TOCOM-B","RSS3","FUTURE","24Jun19","LAST")
=RTD("XTAPI.RTD","","Instr","SGX-B","TF","FUTURE","Jan20","ASK")
=RTD("XTAPI.RTD","","Instr","SGX-B","TF","FUTURE","Feb19","ASK")
=RTD("XTAPI.RTD","","Instr","SGX-B","TF","FUTURE","Mar19","ASK")
=RTD("XTAPI.RTD","","Instr","SGX-B","TF","SPREAD","Calendar: 1xTF Apr19:-1xTF Oct19","ASK")

"""
