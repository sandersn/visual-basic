'InStrtst.bas
'InStr test program
'Notes: 1.start position is 1 not 0.
'   2.calling instr a second time with the position returned from a previous
'call returns the same position unless you add one to the returned position.
'   3.trim will cut "  " down to "" (note: ltrim or rtrim will also do it,
'and probably faster.)
'   4. mid works with a length of 0((pos1 + 1) - pos2)
DIM s1 AS STRING
DIM s2 AS STRING
DIM s3 AS STRING
DIM s4 AS STRING
DIM pos1 AS INTEGER
DIM pos2 AS INTEGER
CLS

'set up some sample strings with example script lines in them
s1 = "j000,999"
s2 = "s,000,ZZ ZZ ZZ"
s3 = "p,00,11,+2,-3,ZZZZ"
s4 = "Thread:000,999"

'test first string
PRINT
PRINT "First string: " + s1
pos1 = INSTR(s1, ",")
PRINT MID$(s1, 2, pos1 - 2)
PRINT RIGHT$(s1, LEN(s1) - pos1)

'test second string
PRINT
PRINT "Second string: " + s2
pos1 = INSTR(s2, ",")
pos2 = INSTR(pos1 + 1, s2, ",")
PRINT MID$(s2, pos1 + 1, pos2 - (pos1 + 1))
PRINT RIGHT$(s2, LEN(s2) - pos2)

'test third string
PRINT
PRINT "Third string: " + s3
pos1 = INSTR(s3, ",")
pos2 = INSTR(pos1 + 1, s3, ",")
temp$ = MID$(s3, pos1 + 1, pos2 - (pos1 + 1))
PRINT temp$
pos1 = pos2
pos2 = INSTR(pos1 + 1, s3, ",")
temp$ = MID$(s3, pos1 + 1, pos2 - (pos1 + 1))
PRINT temp$
pos1 = pos2
pos2 = INSTR(pos1 + 1, s3, ",")
temp$ = MID$(s3, pos1 + 1, pos2 - (pos1 + 1))
PRINT temp$
pos1 = pos2
pos2 = INSTR(pos1 + 1, s3, ",")
temp$ = MID$(s3, pos1 + 1, pos2 - (pos1 + 1))
PRINT temp$
pos1 = pos2
temp$ = RIGHT$(s3, LEN(s3) - (pos2))
PRINT temp$

'test the fourth string
PRINT
PRINT "Fourth string: " + s4
pos1 = INSTR(s4, ":")
pos2 = INSTR(s4, ",")
PRINT LEFT$(s4, pos1 - 1)
PRINT MID$(s4, pos1 + 1, pos2 - (pos1 + 1))
PRINT RIGHT$(s4, LEN(s4) - pos2)

