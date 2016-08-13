set safe off

lcmcod = '0105058'

lcpfile = 'inlpu\p' + lcmcod + '.dbf'
lcafile = 'qareg\a' + lcmcod + '.dbf'
select 2
use (lcafile)
select 3
use (lcpfile)

sele recid,s_pol,n_pol from (lcpfile) where recid in (sele val(recid) from (lcafile) where ans_r='0*0');
	and n_pol>99999 into dbf add

select 0
use d:\prg\ikarsmo\qareg\qreglpu
zap
appe from add

quit