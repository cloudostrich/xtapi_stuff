from win32com.client.gencache import EnsureDispatch



def sub_sprdEOD(pinstr, exch, prod, leg1, leg2):
    pinstr.Exchange = exch
    pinstr.Product = prod
    pinstr.ProdType = 'SPREAD'
    pinstr.Contract = f'Calendar: 1xTF {leg1}:-1xTF {leg2}'

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

# subsribe spreads
#=RTD("XTAPI.RTD","","Instr","SGX-B","TF","SPREAD","Calendar: 1xTF Aug19:-1xTF Nov19","ASK")
sub_sprdEOD(h2, 'SGX-B', 'TF', 'Mar19', 'May19')
sub_sprdEOD(h3, 'SGX-B', 'TF', 'Mar19', 'Jun19')
sub_sprdEOD(h4, 'SGX-B', 'TF', 'Mar19', 'Jul19')
sub_sprdEOD(h5, 'SGX-B', 'TF', 'Mar19', 'Aug19')

sub_sprdEOD(j2, 'SGX-B', 'TF', 'Apr19', 'Jun19')
sub_sprdEOD(j3, 'SGX-B', 'TF', 'Apr19', 'Jul19')
sub_sprdEOD(j4, 'SGX-B', 'TF', 'Apr19', 'Aug19')
sub_sprdEOD(j5, 'SGX-B', 'TF', 'Apr19', 'Sep19')

sub_sprdEOD(k2, 'SGX-B', 'TF', 'May19', 'Jul19')
sub_sprdEOD(k3, 'SGX-B', 'TF', 'May19', 'Aug19')
sub_sprdEOD(k4, 'SGX-B', 'TF', 'May19', 'Sep19')
sub_sprdEOD(k5, 'SGX-B', 'TF', 'May19', 'Oct19')

sub_sprdEOD(m2, 'SGX-B', 'TF', 'Jun19', 'Aug19')
sub_sprdEOD(m3, 'SGX-B', 'TF', 'Jun19', 'Sep19')
sub_sprdEOD(m4, 'SGX-B', 'TF', 'Jun19', 'Oct19')
sub_sprdEOD(m5, 'SGX-B', 'TF', 'Jun19', 'Nov19')

sub_sprdEOD(n2, 'SGX-B', 'TF', 'Jul19', 'Sep19')
sub_sprdEOD(n3, 'SGX-B', 'TF', 'Jul19', 'Oct19')
sub_sprdEOD(n4, 'SGX-B', 'TF', 'Jul19', 'Nov19')
sub_sprdEOD(n5, 'SGX-B', 'TF', 'Jul19', 'Dec19')

sub_sprdEOD(q2, 'SGX-B', 'TF', 'Aug19', 'Oct19')
sub_sprdEOD(q3, 'SGX-B', 'TF', 'Aug19', 'Nov19')
sub_sprdEOD(q4, 'SGX-B', 'TF', 'Aug19', 'Dec19')

sub_sprdEOD(u2, 'SGX-B', 'TF', 'Sep19', 'Nov19')
sub_sprdEOD(u3, 'SGX-B', 'TF', 'Sep19', 'Dec19')

sub_sprdEOD(v2, 'SGX-B', 'TF', 'Oct19', 'Dec19')

sub_sprdEOD(x2, 'SGX-B', 'TF', 'Nov19', 'Jan20')
