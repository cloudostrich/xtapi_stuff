import pythoncom
from time import sleep
from win32com.client import Dispatch, DispatchWithEvents, getevents
from win32com.client.gencache import EnsureDispatch, EnsureModule
import osbrain

Gate = None
Notify = None

class InstrNotify():#getevents('XTAPI.TTInstrNotify')):
    def __init__(self):
##        ns = run_nameserver()
        self.alice = osbrain.run_agent('Alice')
##    bob = run_agent('Bob')
        self.addr = self.alice.bind('PUSH', alias='main')
##    bob.connect(addr, handler=log_message)

    def Subscribe(self, pInstr):
        self.AttachInstrument(pInstr)
        pInstr.Open(0)
        print('subscribed...')

##    def AttachInstrument(self, pInstr):
##        self.AttachInstrument(pInstr)
##        pInstr.Open(0)

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

        upd = f'[UPDATE] {bidqty} | {bid} / {ask} | {askqty} :: {last} | {lastqty}'
        self.alice.send('main', upd)

##        print('[UPDATE] %s: %s/%s' % (contract, bid, ask))


def Connect():
    global Notify, Gate
    #the below is required in order to establish the com-object links
    #that way you don't need to run makepy first
    EnsureModule('{98B8AE14-466F-11D6-A27B-00B0D0F3CCA6}', 0, 1, 0)

    Gate = EnsureDispatch('XTAPI.TTGate')
    Notify = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
##    NOTIFY = Dispatch('XTAPI.TTInstrNotify', InstrNotify)
def log_message(agent, message):
    agent.log_info(f'From XTapi: {message}')
    
def qq():
    Gate.XTAPITerminate()
    ns.shutdown()
    
def main():
    ns = osbrain.run_nameserver()
##    alice = run_agent('Alice')
    bob = osbrain.run_agent('Bob')

    # System configuration
##    addr = alice.bind('PUSH', alias='main')
    
    
    Connect()
    bob.connect(Notify.addr, handler=log_message)
    pInstr = EnsureDispatch('XTAPI.TTInstrObj')
    pInstr.Exchange = 'TOCOM-B'
    pInstr.Product  = 'RSS3'
    pInstr.Contract = '24Jun19'
    pInstr.ProdType = 'FUTURE'

    Notify.Subscribe(pInstr)


    for i in range(10):
        print('pumping...')
        pythoncom.PumpWaitingMessages()
        sleep(1.0)
		
"""
Found instrument:
--> Contract: CL Mar13
--> Exchange: CME-A
[UPDATE] CL Mar13: 9760/9764
=RTD("XTAPI.RTD","","Instr","TOCOM-B","RSS3","FUTURE","24Jun19","LAST")
"""
