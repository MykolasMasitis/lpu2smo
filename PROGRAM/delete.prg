set safe off

lcmcod = '0201047'

lcsfile = 'inlpu\s' + lcmcod + '.dbf'
lcufile = 'posmed\u' + lcmcod + '.dbf'

select 1
use (lcsfile)

select 2
use (lcufile)
index on recid tag recid

select 1
set rela to recid into 2

repl a.k_u with a.k_u-b.k_u all
set filt to k_u < 1
delete all
pack

quit