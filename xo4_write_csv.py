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
        upd = {'time':pd.datetime.now(), 'bidqty': bidqty, 'bid': bid, 'ask': ask,
               'askqty': askqty, 'last': last, 'lastqty': lastqty}
        self.agent_notify.send(self.addr_alias, [upd, contract])
        
class WriteDisk(osbrain.Agent):
    # add self.store, self.cache
    def on_init(self):
        self.CACHE = {}
        self.STORE = f"db_{self.get_attr('name')}.csv"
        self.lst = None

    def custom_log(self, message):
        self.log_info(f'Got it: {message}')        

    def process_row(self, m=[None, None], max_len=10): ##    def process_row(self, d, key, max_len=50): #, _cache=self.CACHE):
        """Creates a dict with key holding a list of dicts of tick data
        Append row d to the store 'key'.
        When the number of items in the key's cache reaches max_len,
        append the list of rows to the HDF5 store and clear the list."""
        # keep the rows for each key separate.
        self.lst = self.CACHE.setdefault(m[1], []) #key, []) #set default key for dict CACHE
        if len(self.lst) >= max_len:
            self.store_and_clear(self.lst, m[1]) #key)
        self.lst.append(m[0])

    def store_and_clear(self, lst, key):
        """
        Convert key's cache list to a DataFrame and append that to HDF5.
        """
        df = pd.DataFrame(lst)
##        df.set_index(['time'], inplace = True)
##        with pd.HDFStore(self.STORE) as store:
##            store.append(key, df)
        try:
            with open(self.STORE,'a') as store:
                df.to_csv(store, mode='a', header=False)
        except FileNotFoundError:
            with open(self.STORE,'w') as store:
                df.to_csv(self.STORE)
        print(f'Wrote to disk: {key}, {lst[0]}')
        self.lst.clear()
        

    def get_latest(self, key):#, _cache=self.CACHE): qq(ns,Gate)
        self.store_and_clear(self.CACHE[key], key)
##        with pd.HDFStore(self.STORE) as store:
##            return store[key]
        
class Greeter(osbrain.Agent):
    def on_init(self):
        self.bind('PUSH', alias='main')

    def hello(self, name):
        self.send('main', f'Hello Asshole {name}')
        
def Connect():
    global NotifyTF, NotifySprd, NotifyTcm, Gate
    #the below is required in order to establish the com-object links
    #that way you don't need to run makepy first
    EnsureModule('{98B8AE14-466F-11D6-A27B-00B0D0F3CCA6}', 0, 1, 0)

    Gate = EnsureDispatch('XTAPI.TTGate')
    NotifyTF = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    NotifyTF.gen_agent('agentTF','channelTF')
    print('Connected TF...')
    NotifySprd = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    NotifySprd.gen_agent('agentSPRD','channelSPRD')
    print('Connected Spreads...')
    NotifyTcm = DispatchWithEvents('XTAPI.TTInstrNotify', InstrNotify)
    NotifyTcm.gen_agent('agentTcm','channelTcm')
    print('Connected Tocom...')

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
    tcmw = osbrain.run_agent('Tocom-w', base=WriteDisk)
    tfw = osbrain.run_agent('TF-w', base=WriteDisk)
    sprdw = osbrain.run_agent('SPRD-w', base=WriteDisk)
    Connect()
    tcmw.connect(NotifyTcm.addr, handler=['process_row','custom_log'])
    tfw.connect(NotifyTF.addr, handler=['process_row','custom_log'])
    sprdw.connect(NotifySprd.addr, handler=['process_row','custom_log'])

    ''' init tocom rss
    =RTD("XTAPI.RTD","","Instr","TOCOM-B","RSS3","FUTURE","25Jan19","ASK")
    =RTD("XTAPI.RTD","","Instr","SGX-B","TF","FUTURE","Feb19","ASK")
    '''
    tcm_25Jul19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_22Feb19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_25Mar19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_22Apr19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_27May19 = EnsureDispatch('XTAPI.TTInstrObj')
    tcm_24Jun19 = EnsureDispatch('XTAPI.TTInstrObj')
    print('TOCOM dispatched...')

    tcm_lst = (tcm_25Jul19, tcm_22Feb19, tcm_25Mar19,
               tcm_22Apr19, tcm_27May19, tcm_24Jun19)
    tcm_mth = ('25Jul19', '22Feb19', '25Mar19', '22Apr19', '27May19', '24Jun19')
        
    ''' dispatch tocom rss
    sub_fut(tcm_24Jun19, 'TOCOM-B', 'RSS3', '24Jun19',NotifyTcm) '''
    
    for l, m in zip(tcm_lst, tcm_mth):
        sub_fut(l, 'TOCOM-B', 'RSS3', m,NotifyTcm)
        
    NotifyTcm.UpdateFilter = 'Bid, Ask, Last, LastQty'

    # dispatch and subscribe TF
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
    print('TF dispatched...')

    tf_lst = (mar19, apr19, may19, jun19, jul19, aug19, sep19, oct19, nov19, dec19, jan20)
    
    tf_mth = ('Mar19', 'Apr19', 'May19', 'Jun19', 'Jul19', 'Aug19', 'Sep19', 'Oct19', 'Nov19', 'Dec19','Jan20')

    for l, m in zip(tf_lst, tf_mth):
        sub_fut(l, 'SGX-B', 'TF', m, NotifyTF)
        
    NotifyTF.UpdateFilter = 'Bid, Ask, Last, LastQty'

    # Dispatch TF spreads
    h2 = EnsureDispatch('XTAPI.TTInstrObj')
    h3 = EnsureDispatch('XTAPI.TTInstrObj')
    h4 = EnsureDispatch('XTAPI.TTInstrObj')
    h5 = EnsureDispatch('XTAPI.TTInstrObj')

    j2 = EnsureDispatch('XTAPI.TTInstrObj')
    j3 = EnsureDispatch('XTAPI.TTInstrObj')
    j4 = EnsureDispatch('XTAPI.TTInstrObj')
    j5 = EnsureDispatch('XTAPI.TTInstrObj')

    k2 = EnsureDispatch('XTAPI.TTInstrObj')
    k3 = EnsureDispatch('XTAPI.TTInstrObj')
    k4 = EnsureDispatch('XTAPI.TTInstrObj')
    k5 = EnsureDispatch('XTAPI.TTInstrObj')

    m2 = EnsureDispatch('XTAPI.TTInstrObj')
    m3 = EnsureDispatch('XTAPI.TTInstrObj')
    m4 = EnsureDispatch('XTAPI.TTInstrObj')
    m5 = EnsureDispatch('XTAPI.TTInstrObj')

    n2 = EnsureDispatch('XTAPI.TTInstrObj')
    n3 = EnsureDispatch('XTAPI.TTInstrObj')
    n4 = EnsureDispatch('XTAPI.TTInstrObj')
    n5 = EnsureDispatch('XTAPI.TTInstrObj')

    q2 = EnsureDispatch('XTAPI.TTInstrObj')
    q3 = EnsureDispatch('XTAPI.TTInstrObj')
    q4 = EnsureDispatch('XTAPI.TTInstrObj')
##    q5 = EnsureDispatch('XTAPI.TTInstrObj')

    u2 = EnsureDispatch('XTAPI.TTInstrObj')
    u3 = EnsureDispatch('XTAPI.TTInstrObj')
##    u4 = EnsureDispatch('XTAPI.TTInstrObj')
##    u5 = EnsureDispatch('XTAPI.TTInstrObj')

    v2 = EnsureDispatch('XTAPI.TTInstrObj')

    x2 = EnsureDispatch('XTAPI.TTInstrObj')
    print('Spreads dispatched...')

    ''' subsribe spreads
    =RTD("XTAPI.RTD","","Instr","SGX-B","TF","SPREAD","Calendar: 1xTF Aug19:-1xTF Nov19","ASK")'''
    
    sub_sprd(h2, 'SGX-B', 'TF', 'Mar19', 'May19', NotifySprd)
    sub_sprd(h3, 'SGX-B', 'TF', 'Mar19', 'Jun19', NotifySprd)
    sub_sprd(h4, 'SGX-B', 'TF', 'Mar19', 'Jul19', NotifySprd)
    sub_sprd(h5, 'SGX-B', 'TF', 'Mar19', 'Aug19', NotifySprd)
    
    sub_sprd(j2, 'SGX-B', 'TF', 'Apr19', 'Jun19', NotifySprd)
    sub_sprd(j3, 'SGX-B', 'TF', 'Apr19', 'Jul19', NotifySprd)
    sub_sprd(j4, 'SGX-B', 'TF', 'Apr19', 'Aug19', NotifySprd)
    sub_sprd(j5, 'SGX-B', 'TF', 'Apr19', 'Sep19', NotifySprd)

    sub_sprd(k2, 'SGX-B', 'TF', 'May19', 'Jul19', NotifySprd)
    sub_sprd(k3, 'SGX-B', 'TF', 'May19', 'Aug19', NotifySprd)
    sub_sprd(k4, 'SGX-B', 'TF', 'May19', 'Sep19', NotifySprd)
    sub_sprd(k5, 'SGX-B', 'TF', 'May19', 'Oct19', NotifySprd)

    sub_sprd(m2, 'SGX-B', 'TF', 'Jun19', 'Aug19', NotifySprd)
    sub_sprd(m3, 'SGX-B', 'TF', 'Jun19', 'Sep19', NotifySprd)
    sub_sprd(m4, 'SGX-B', 'TF', 'Jun19', 'Oct19', NotifySprd)
    sub_sprd(m5, 'SGX-B', 'TF', 'Jun19', 'Nov19', NotifySprd)

    sub_sprd(n2, 'SGX-B', 'TF', 'Jul19', 'Sep19', NotifySprd)
    sub_sprd(n3, 'SGX-B', 'TF', 'Jul19', 'Oct19', NotifySprd)
    sub_sprd(n4, 'SGX-B', 'TF', 'Jul19', 'Nov19', NotifySprd)
    sub_sprd(n5, 'SGX-B', 'TF', 'Jul19', 'Dec19', NotifySprd)

    sub_sprd(q2, 'SGX-B', 'TF', 'Aug19', 'Oct19', NotifySprd)
    sub_sprd(q3, 'SGX-B', 'TF', 'Aug19', 'Nov19', NotifySprd)
    sub_sprd(q4, 'SGX-B', 'TF', 'Aug19', 'Dec19', NotifySprd)

    sub_sprd(u2, 'SGX-B', 'TF', 'Sep19', 'Nov19', NotifySprd)
    sub_sprd(u3, 'SGX-B', 'TF', 'Sep19', 'Dec19', NotifySprd)

    sub_sprd(v2, 'SGX-B', 'TF', 'Oct19', 'Dec19', NotifySprd)

    sub_sprd(x2, 'SGX-B', 'TF', 'Nov19', 'Jan20', NotifySprd)

    NotifySprd.UpdateFilter = 'Bid, Ask, Last, LastQty'

    for i in range(15):
        print('pumping...')
        pythoncom.PumpWaitingMessages()
        sleep(1.0)
