'Epsilon Algorithm, Created by Simon Johnson
'Uses storing encrypted passwd's, or producing message digests.
'

Public Function Hash(byval text as string) as string
a=1
For i = 1 to len(text)
    a=sqr(a*i*asc(mid(text,i,1))) 'Numeric Hash
Next i
Rnd(-1)
Randomize a 'seed PRNG

For i = 1 to 16
    Hash = Hash & Chr(int(rnd*256))
Next i
End function