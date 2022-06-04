CONST CPPBeginTag = "//"
DIM CPPEndTag AS STRING
CPPEndTag = CHR$(10) + CHR$(13)
CONST CBeginTag = "/*"
CONST CEndTag = "*/"
DIM strFilename AS STRING
DIM intFileno AS INTEGER
DIM ch AS STRING
DIM readText AS STRING
DIM bReading AS INTEGER
DIM endTag AS STRING
DIM pAddr AS LONG
DIM textAddr AS LONG
'read filename from command line
strFilename = COMMAND$
ch = "  "

'IF Dir$(strFilename) = "" THEN
    'PRINT "File not found. Retype filename"
    'END
'ELSE
'open file in binary mode
    intFileno = FREEFILE
    pAddr = 1
    OPEN strFilename FOR BINARY AS #intFileno
    DO

        GET #intFileno, pAddr, ch
        IF bReading = 1 THEN
'read all text in until find a newl or */
            IF ch = endTag THEN
'Ucase(text found)
                readText = UCASE$(readText)
'write back to file
                textLen = LEN(readText) - 1
                PUT #intFileno, textAddr + 2, textLen
                bReading = 0
            ELSE
                readText = readText + RIGHT$(ch, 1)
            END IF
        ELSE
'loop through each address of the file looking for // or /*
            IF ch = CBeginTag OR ch = CPPBeginTag THEN
                IF ch = CBeginTag THEN
                    endTag = CBeginTag
                ELSE
                    endTag = CPPEndTag
                END IF
                textAddr = pAddr
                readText = ""   'reinit readText
                bReading = 1
            END IF
        END IF
        pAddr = pAddr + 1
    LOOP UNTIL EOF(intFileno)
    CLOSE intFileno
    PRINT "Process. Now be happy"
'END IF

   




'input:filename
'output:confirmation to user/maybe show contents of updated file

