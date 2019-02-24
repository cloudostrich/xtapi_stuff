import pythoncom
from time import sleep
from win32com.client import Dispatch, DispatchWithEvents, getevents
from win32com.client.gencache import EnsureDispatch, EnsureModule
import osbrain

Gate = None
NotifyTF = None
ns = None

class InstrNotify():#getevents('XTAPI.TTInstrNotify')):
    def __init__(self):
        self.alice = osbrain.run_agent('Alice')
        self.addr = self.alice.bind('PUSH', alias='main')

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

##        self.UpdateFilter = 'Bid', 'Ask', 'Last', 'LastQty'
##        upd = f'[UPDATE] {contract}: {bidqty} | {bid} / {ask} | {askqty} :: {last} | {lastqty}'
        upd = {'contract':contract, 'bidqty': bidqty, 'bid': bid, 'ask': ask, 'askqty': askqty, 'last': last, 'lastqty': lastqty}
        self.alice.send('main', upd)
##        print('[UPDATE] %s: %s/%s' % (contract, bid, ask))


def Connect():
    global NotifyTF, Gate
    #the below is required in order to establish the com-object links
    #that way you don't need to run makepy first
    EnsureModule('{98B8AE14-466F-11D6-A27B-00B0D0F3CCA6}', 0, 1, 0)

    Gate = EnsureDispatch('XTAPI.TTGate')
    NotifyTF = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    print('Connected...')
##    NOTIFY = Dispatch('XTAPI.TTInstrNotify', InstrNotify)
    
def log_message(agent, message):
    agent.log_info(f'From XTapi: {message}')
    
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
    pinstr.Contract = "Calendar: 1xTF {leg1}:-1xTF {leg2}"
    notifier.Subscribe(pinstr)
    
def main():
    global ns
    ns = osbrain.run_nameserver()
    bob = osbrain.run_agent('Bob')
    Connect()
    bob.connect(NotifyTF.addr, handler=log_message)
##    pInstr = EnsureDispatch('XTAPI.TTInstrObj')
##    pInstr.Exchange = 'TOCOM-B'
##    pInstr.Product  = 'RSS3'
##    pInstr.Contract = '24Jun19'
##    pInstr.ProdType = 'FUTURE'
##    Notify.Subscribe(pInstr)

    # dispatch tf outrights
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
    jan20 = EnsureDispatch('XTAPI.TTInstrObj')

    # dispatch tf spreads
    

    sub_fut(mar19, 'SGX-B', 'TF', 'Mar19',NotifyTF)
    sub_fut(apr19, 'SGX-B', 'TF', 'Apr19',NotifyTF)
    sub_fut(may19, 'SGX-B', 'TF', 'May19',NotifyTF)
    sub_fut(jun19, 'SGX-B', 'TF', 'Jun19',NotifyTF)
    sub_fut(jul19, 'SGX-B', 'TF', 'Jul19',NotifyTF)
    sub_fut(aug19, 'SGX-B', 'TF', 'Aug19',NotifyTF)
    sub_fut(sep19, 'SGX-B', 'TF', 'Sep19',NotifyTF)
    sub_fut(oct19, 'SGX-B', 'TF', 'Oct19',NotifyTF)
    sub_fut(nov19, 'SGX-B', 'TF', 'Nov19',NotifyTF)
    sub_fut(dec19, 'SGX-B', 'TF', 'Dec19',NotifyTF)
    sub_fut(jan20, 'SGX-B', 'TF', 'Jan20',NotifyTF)
    
    NotifyTF.UpdateFilter = 'Bid, Ask, Last, LastQty'


    for i in range(10):
        print('pumping...')
        pythoncom.PumpWaitingMessages()
        sleep(1.0)


