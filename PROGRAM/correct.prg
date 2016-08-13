set safe off

lcmcod = '8143007'

lcsfile = 'inlpu\s' + lcmcod + '.dbf'
lcufile = 'posmed\s' + lcmcod + '.dbf'

select 1
use (lcsfile)

select 2
use (lcufile)
index on recid tag recid

select 1
set rela to recid into 2

set filt to a.recid=b.recid and !a.cod=b.cod
repl a.cod with b.cod all

set filt to a.recid=b.recid and !a.tip=b.tip
repl a.tip with b.tip all

set filt to a.recid=b.recid and !a.ds=b.ds
repl a.ds with b.ds all

set filt to k_u < 1
delete all
pack

quit