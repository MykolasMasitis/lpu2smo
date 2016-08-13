PROCEDURE HVScreen

nW = 0
nH = 0

=ScreenSize(@nW, @nH)

MESSAGEBOX('–¿«–≈ÿ≈Õ»≈ ÃŒÕ»“Œ–¿: '+ALLTRIM(STR(nW))+'x'+ALLTRIM(STR(nH)),0+64,'')


RETURN 