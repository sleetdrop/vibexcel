# Sudoku Named LAMBDA Snapshot

Generated file — do not edit manually. The Excel workbook is the canonical executable artifact; formulas are normalized for display by removing storage-only `_xlfn.` and `_xlpm.` prefixes.

- Workbook: `sudoku.xlsx`
- Workbook version: `1.0.0`
- Workbook SHA-256: `9ed530cd79df2fb73c90a1441676cdf4548df8fbd217d500bb17fc6dbf4316c3`
- Extraction date: `2026-07-28`
- Formula count: `51`

## `SDK_BasicHint`

```excel
=LAMBDA(board,LET(a,SDK_HintNaked(board),b,SDK_HintHiddenRow(board),c,SDK_HintHiddenCol(board),d,SDK_HintHiddenBox(board),IF(a<>"",a,IF(b<>"",b,IF(c<>"",c,d)))))
```

## `SDK_CandidateMatrix`

```excel
=LAMBDA(board,MAKEARRAY(9,9,LAMBDA(r,c,SDK_Candidates(board,r,c))))
```

## `SDK_Candidates`

Returns legal digits for one square using row, column, and box exclusion; preserves blank arrays to avoid empty-unit #CALC errors.

```excel
=LAMBDA(board,r,c,LET(cur,INDEX(board,r,c),IF(LEN(cur&"")>0,"",LET(rowVals,INDEX(board,r,0),colVals,TOCOL(INDEX(board,0,c),0),br,1+3*QUOTIENT(r-1,3),bc,1+3*QUOTIENT(c-1,3),boxVals,TOCOL(INDEX(board,SEQUENCE(3,,br),SEQUENCE(,3,bc)),0),used,TOCOL(VSTACK(rowVals,colVals,boxVals),0),TEXTJOIN("",TRUE,_xlws.FILTER(SEQUENCE(9),ISNA(XMATCH(SEQUENCE(9),used)),""))))))
```

## `SDK_CoachHint`

```excel
=LAMBDA(board,SDK_CoachHintV2(board))
```

## `SDK_CoachHintV2`

```excel
=LAMBDA(board,LET(basicHint,SDK_BasicHint(board),pointingHint,SDK_PointingHint(board),claimingHint,SDK_HintClaiming(board),pairHint,SDK_HintNakedPair(board),IF(basicHint<>"",basicHint,IF(pointingHint<>"",pointingHint,IF(claimingHint<>"",claimingHint,IF(pairHint<>"",pairHint,"No supported logical step is currently available."))))))
```

## `SDK_CoachTrace`

```excel
=LAMBDA(board,SDK_CoachTraceV2(board))
```

## `SDK_CoachTraceV2`

```excel
=LAMBDA(board,LET(hintText,SDK_CoachHintV2(board),parts,TEXTSPLIT(hintText," · "),partCount,COLUMNS(parts),lineOne,IF(partCount>1,"Technique — "&INDEX(parts,1,1)&" | Pattern — "&INDEX(parts,1,2),"Status — "&INDEX(parts,1,1)),lineTwo,IF(partCount>2,"Action — "&TEXTJOIN(" · ",TRUE,DROP(parts,,2)),"Action — No further supported elimination."),VSTACK(lineOne,lineTwo)))
```

## `SDK_Conflict`

```excel
=LAMBDA(board,r,c,LET(v,INDEX(board,r,c),IF(v="",FALSE,LET(br,1+3*QUOTIENT(r-1,3),bc,1+3*QUOTIENT(c-1,3),OR(SUM(--(INDEX(board,r,0)=v))>1,SUM(--(INDEX(board,0,c)=v))>1,SUM(--(INDEX(board,SEQUENCE(3,,br),SEQUENCE(,3,bc))=v))>1)))))
```

## `SDK_CoordRole`

```excel
=LAMBDA(hintText,coord,LET(removePos,IFERROR(SEARCH("Remove",hintText),1000000),coordPos,IFERROR(SEARCH(coord,hintText),0),IF(coordPos=0,"",IF(coordPos>removePos,"TARGET","SOURCE"))))
```

## `SDK_CountSolutions`

```excel
=LAMBDA(board,limit,LET(conflicts,SUM(MAKEARRAY(9,9,LAMBDA(r,c,--SDK_Conflict(board,r,c)))),dead,SDK_HasDeadEnd(board),emptyCount,SUM(--(LEN(board&"")=0)),IF(OR(limit<=0,conflicts>0,dead),0,IF(emptyCount=0,1,LET(candidates,SDK_CandidateMatrix(board),lengths,IF(LEN(board&"")=0,LEN(candidates),99),minLength,MIN(lengths),position,XMATCH(minLength,TOCOL(lengths)),r,QUOTIENT(position-1,9)+1,c,MOD(position-1,9)+1,options,INDEX(candidates,r,c)&"",SDK_CountTry(board,r,c,options,1,limit))))))
```

## `SDK_CountTry`

```excel
=LAMBDA(board,r,c,options,i,limit,IF(OR(limit<=0,i>LEN(options)),0,LET(nextBoard,SDK_SetCell(board,r,c,--MID(options,i,1)),first,SDK_CountSolutions(nextBoard,limit),IF(first>=limit,limit,first+SDK_CountTry(board,r,c,options,i+1,limit-first)))))
```

## `SDK_DeadBox`

Returns the first missing box digit with no candidate position.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(b,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,br,1+3*QUOTIENT(b-1,3),bc,1+3*MOD(b-1,3),bv,INDEX(board,SEQUENCE(3,,br),SEQUENCE(,3,bc)),cv,INDEX(c,SEQUENCE(3,,br),SEQUENCE(,3,bc)),absent,SUM(--(bv=d))=0,places,SUM(--ISNUMBER(SEARCH(d&"",cv&""))),IF(AND(absent,places=0),"Dead end · Box "&b&" has no place for digit "&d&".","")))))))
```

## `SDK_DeadCell`

Returns the first empty square with no legal candidates.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),m,(board="")*(c=""),p,IFERROR(XMATCH(1,TOCOL(--m)),0),IF(p=0,"",LET(r,QUOTIENT(p-1,9)+1,k,MOD(p-1,9)+1,"Dead end · R"&r&"C"&k&" has no legal candidates."))))
```

## `SDK_DeadCol`

Returns the first missing column digit with no candidate position; array-safe for recursive solver boards.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(k,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,colBoard,INDEX(board,0,k),absent,SUM(--(colBoard=d))=0,places,SUM(--ISNUMBER(SEARCH(d&"",INDEX(c,0,k)&""))),IF(AND(absent,places=0),"Dead end · Column "&k&" has no place for digit "&d&".","")))))))
```

## `SDK_DeadEnd`

Returns the first provable local contradiction in priority order.

```excel
=LAMBDA(board,LET(a,SDK_DeadCell(board),b,SDK_ForcedRow(board),c,SDK_ForcedCol(board),d,SDK_ForcedBox(board),e,SDK_DeadRow(board),f,SDK_DeadCol(board),g,SDK_DeadBox(board),IF(a<>"",a,IF(b<>"",b,IF(c<>"",c,IF(d<>"",d,IF(e<>"",e,IF(f<>"",f,g))))))))
```

## `SDK_DeadRow`

Returns the first missing row digit with no candidate position; array-safe for recursive solver boards.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(r,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,rowBoard,INDEX(board,r,0),absent,SUM(--(rowBoard=d))=0,places,SUM(--ISNUMBER(SEARCH(d&"",INDEX(c,r,0)&""))),IF(AND(absent,places=0),"Dead end · Row "&r&" has no place for digit "&d&".","")))))))
```

## `SDK_FallbackReveal`

```excel
=LAMBDA(board,solution,LET(p,IFERROR(XMATCH("",TOCOL(board)),0),IF(p=0,"Puzzle complete.",LET(r,QUOTIENT(p-1,9)+1,c,MOD(p-1,9)+1,d,INDEX(solution,r,c),"Coach reveal · R"&r&"C"&c&" = "&d&" · No supported logical deduction is available; this value comes from the certified solution."))))
```

## `SDK_ForcedBox`

Returns the first box containing duplicate forced singleton candidates.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(b,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,br,1+3*QUOTIENT(b-1,3),bc,1+3*MOD(b-1,3),v,TOCOL(INDEX(c,SEQUENCE(3,,br),SEQUENCE(,3,bc)))&"",n,SUM(--(v=d&"")),pos,_xlws.FILTER(SEQUENCE(9),v=d&"",""),rr,br+QUOTIENT(pos-1,3),kk,bc+MOD(pos-1,3),IF(n>1,"Dead end · Box "&b&" forces digit "&d&" into "&TEXTJOIN(" and ",TRUE,"R"&rr&"C"&kk)&".","")))))))
```

## `SDK_ForcedCol`

Returns the first column containing duplicate forced singleton candidates.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(k,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,v,INDEX(c,0,k)&"",n,SUM(--(v=d&"")),pos,_xlws.FILTER(SEQUENCE(9),v=d&"",""),IF(n>1,"Dead end · Column "&k&" forces digit "&d&" into "&TEXTJOIN(" and ",TRUE,"R"&pos&"C"&k)&".","")))))))
```

## `SDK_ForcedRow`

Returns the first row containing duplicate forced singleton candidates.

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(r,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,v,INDEX(c,r,0)&"",n,SUM(--(v=d&"")),pos,_xlws.FILTER(SEQUENCE(,9),v=d&"",""),IF(n>1,"Dead end · Row "&r&" forces digit "&d&" into "&TEXTJOIN(" and ",TRUE,"R"&r&"C"&pos)&".","")))))))
```

## `SDK_HasDeadEnd`

TRUE when SDK_DeadEnd returns a contradiction.

```excel
=LAMBDA(board,SDK_DeadEnd(board)<>"")
```

## `SDK_HasSolution`

Experimental Boolean wrapper around SDK_Solve. Not yet connected to live game pages.

```excel
=LAMBDA(board,IFERROR(LET(result,SDK_Solve(board),AND(ROWS(result)=9,COLUMNS(result)=9)),FALSE))
```

## `SDK_HintAssignmentValid`

```excel
=LAMBDA(board,solution,hint,IFERROR(LET(afterR,TEXTAFTER(hint,"R"),r,--TEXTBEFORE(afterR,"C"),afterC,TEXTAFTER(afterR,"C"),c,--TEXTBEFORE(afterC," = "),d,--TEXTBEFORE(TEXTAFTER(afterC," = ")," ·"),AND(LEN(INDEX(board,r,c)&"")=0,INDEX(solution,r,c)=d,ISNUMBER(SEARCH(d&"",SDK_Candidates(board,r,c)&"")))),FALSE))
```

## `SDK_HintClaiming`

```excel
=LAMBDA(board,LET(a,SDK_HintClaimingRow(board),b,SDK_HintClaimingCol(board),IF(a<>"",a,b)))
```

## `SDK_HintClaimingCol`

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(k,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,v,INDEX(c,0,k),m,ISNUMBER(SEARCH(d&"",v&"")),n,SUM(--m),rs,_xlws.FILTER(SEQUENCE(9),m,""),bands,1+QUOTIENT(rs-1,3),one,IF(n>1,MIN(bands)=MAX(bands),FALSE),band,IF(one,INDEX(bands,1,1),1),br,1+3*(band-1),bc,1+3*QUOTIENT(k-1,3),sub,INDEX(c,SEQUENCE(3,,br),SEQUENCE(,3,bc)),rr,MAKEARRAY(3,3,LAMBDA(x,y,br+x-1)),cc,MAKEARRAY(3,3,LAMBDA(x,y,bc+y-1)),tm,IF(one,(cc<>k)*ISNUMBER(SEARCH(d&"",sub&"")),FALSE),targets,_xlws.FILTER("R"&TOCOL(rr)&"C"&TOCOL(cc),TOCOL(tm),""),txt,TEXTJOIN(", ",TRUE,targets),b,3*(band-1)+1+QUOTIENT(k-1,3),IF(txt<>"","Claiming column · Column "&k&" confines digit "&d&" to box "&b&" · Remove "&d&" from "&txt&".","")))))))
```

## `SDK_HintClaimingRow`

```excel
=LAMBDA(board,LET(c,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(acc,i,IF(acc<>"",acc,LET(r,QUOTIENT(i-1,9)+1,d,MOD(i-1,9)+1,v,INDEX(c,r,0),m,ISNUMBER(SEARCH(d&"",v&"")),n,SUM(--m),cs,_xlws.FILTER(SEQUENCE(,9),m,""),bands,1+QUOTIENT(cs-1,3),one,IF(n>1,MIN(bands)=MAX(bands),FALSE),band,IF(one,INDEX(bands,1,1),1),br,1+3*QUOTIENT(r-1,3),bc,1+3*(band-1),sub,INDEX(c,SEQUENCE(3,,br),SEQUENCE(,3,bc)),rr,MAKEARRAY(3,3,LAMBDA(x,y,br+x-1)),cc,MAKEARRAY(3,3,LAMBDA(x,y,bc+y-1)),tm,IF(one,(rr<>r)*ISNUMBER(SEARCH(d&"",sub&"")),FALSE),targets,_xlws.FILTER("R"&TOCOL(rr)&"C"&TOCOL(cc),TOCOL(tm),""),txt,TEXTJOIN(", ",TRUE,targets),b,3*QUOTIENT(r-1,3)+band,IF(txt<>"","Claiming row · Row "&r&" confines digit "&d&" to box "&b&" · Remove "&d&" from "&txt&".","")))))))
```

## `SDK_HintHiddenBox`

```excel
=LAMBDA(board,LET(c,MAKEARRAY(9,9,LAMBDA(r,k,SDK_Candidates(board,r,k))),n,MAKEARRAY(9,9,LAMBDA(b,d,LET(br,1+3*QUOTIENT(b-1,3),bc,1+3*MOD(b-1,3),SUM(--ISNUMBER(SEARCH(d&"",INDEX(c,SEQUENCE(3,,br),SEQUENCE(,3,bc))&"")))))),p,IFERROR(XMATCH(1,TOCOL(n)),0),IF(p=0,"",LET(b,QUOTIENT(p-1,9)+1,d,MOD(p-1,9)+1,br,1+3*QUOTIENT(b-1,3),bc,1+3*MOD(b-1,3),x,TOCOL(INDEX(c,SEQUENCE(3,,br),SEQUENCE(,3,bc)),0,FALSE),q,XMATCH(TRUE,ISNUMBER(SEARCH(d&"",x&""))),r,br+QUOTIENT(q-1,3),k,bc+MOD(q-1,3),"Hidden single in box · R"&r&"C"&k&" = "&d&" · Digit "&d&" appears in only one candidate set in box "&b&"."))))
```

## `SDK_HintHiddenCol`

```excel
=LAMBDA(board,LET(c,MAKEARRAY(9,9,LAMBDA(r,k,SDK_Candidates(board,r,k))),n,MAKEARRAY(9,9,LAMBDA(k,d,SUM(--ISNUMBER(SEARCH(d&"",INDEX(c,0,k)&""))))),p,IFERROR(XMATCH(1,TOCOL(n)),0),IF(p=0,"",LET(k,QUOTIENT(p-1,9)+1,d,MOD(p-1,9)+1,r,XMATCH(TRUE,ISNUMBER(SEARCH(d&"",INDEX(c,0,k)&""))),"Hidden single in column · R"&r&"C"&k&" = "&d&" · Digit "&d&" appears in only one candidate set in column "&k&"."))))
```

## `SDK_HintHiddenRow`

```excel
=LAMBDA(board,LET(c,MAKEARRAY(9,9,LAMBDA(r,k,SDK_Candidates(board,r,k))),n,MAKEARRAY(9,9,LAMBDA(r,d,SUM(--ISNUMBER(SEARCH(d&"",INDEX(c,r,0)&""))))),p,IFERROR(XMATCH(1,TOCOL(n)),0),IF(p=0,"",LET(r,QUOTIENT(p-1,9)+1,d,MOD(p-1,9)+1,k,XMATCH(TRUE,ISNUMBER(SEARCH(d&"",INDEX(c,r,0)&""))),"Hidden single in row · R"&r&"C"&k&" = "&d&" · Digit "&d&" appears in only one candidate set in row "&r&"."))))
```

## `SDK_HintNaked`

```excel
=LAMBDA(board,LET(c,MAKEARRAY(9,9,LAMBDA(r,k,SDK_Candidates(board,r,k))),m,LEN(c)=1,p,IFERROR(XMATCH(TRUE,TOCOL(m)),0),IF(p=0,"",LET(r,QUOTIENT(p-1,9)+1,k,MOD(p-1,9)+1,d,INDEX(c,r,k),"Naked single · R"&r&"C"&k&" = "&d&" · This is the only legal candidate remaining in the square."))))
```

## `SDK_HintNakedPair`

```excel
=LAMBDA(board,LET(rowHint,SDK_HintNakedPairRow(board),colHint,SDK_HintNakedPairCol(board),boxHint,SDK_HintNakedPairBox(board),IF(rowHint<>"",rowHint,IF(colHint<>"",colHint,boxHint))))
```

## `SDK_HintNakedPairBox`

```excel
=LAMBDA(board,LET(candMatrix,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(foundHint,scanIndex,IF(foundHint<>"",foundHint,LET(rowNum,QUOTIENT(scanIndex-1,9)+1,colNum,MOD(scanIndex-1,9)+1,pairText,INDEX(candMatrix,rowNum,colNum)&"",boxStartRow,1+3*QUOTIENT(rowNum-1,3),boxStartCol,1+3*QUOTIENT(colNum-1,3),unitVals,TOCOL(INDEX(candMatrix,SEQUENCE(3,,boxStartRow),SEQUENCE(,3,boxStartCol)))&"",pairFlag,AND(LEN(pairText)=2,SUM(--(unitVals=pairText))=2),pairPositions,_xlws.FILTER(SEQUENCE(9),unitVals=pairText,""),firstIndex,IF(pairFlag,INDEX(pairPositions,1),0),secondIndex,IF(pairFlag,INDEX(pairPositions,ROWS(pairPositions)),0),rowCoords,boxStartRow+QUOTIENT(SEQUENCE(9)-1,3),colCoords,boxStartCol+MOD(SEQUENCE(9)-1,3),otherMask,(SEQUENCE(9)<>firstIndex)*(SEQUENCE(9)<>secondIndex),hitMask,IF(pairFlag,otherMask*((ISNUMBER(SEARCH(LEFT(pairText,1),unitVals))+ISNUMBER(SEARCH(RIGHT(pairText,1),unitVals)))>0),SEQUENCE(9)=0),targetCells,_xlws.FILTER("R"&rowCoords&"C"&colCoords,hitMask,""),targetText,TEXTJOIN(", ",TRUE,targetCells),firstRow,boxStartRow+QUOTIENT(firstIndex-1,3),firstCol,boxStartCol+MOD(firstIndex-1,3),secondRow,boxStartRow+QUOTIENT(secondIndex-1,3),secondCol,boxStartCol+MOD(secondIndex-1,3),boxNum,1+3*QUOTIENT(rowNum-1,3)+QUOTIENT(colNum-1,3),IF(targetText<>"","Naked pair in box · R"&firstRow&"C"&firstCol&" and R"&secondRow&"C"&secondCol&" contain only "&LEFT(pairText,1)&"/"&RIGHT(pairText,1)&" in box "&boxNum&" · Remove "&LEFT(pairText,1)&" and "&RIGHT(pairText,1)&" from "&targetText&".","")))))))
```

## `SDK_HintNakedPairCol`

```excel
=LAMBDA(board,LET(candMatrix,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(foundHint,scanIndex,IF(foundHint<>"",foundHint,LET(colNum,QUOTIENT(scanIndex-1,9)+1,rowNum,MOD(scanIndex-1,9)+1,pairText,INDEX(candMatrix,rowNum,colNum)&"",unitVals,INDEX(candMatrix,0,colNum)&"",pairFlag,AND(LEN(pairText)=2,SUM(--(unitVals=pairText))=2),pairPositions,_xlws.FILTER(SEQUENCE(9),unitVals=pairText,""),firstPos,IF(pairFlag,INDEX(pairPositions,1),0),secondPos,IF(pairFlag,INDEX(pairPositions,ROWS(pairPositions)),0),otherMask,(SEQUENCE(9)<>firstPos)*(SEQUENCE(9)<>secondPos),hitMask,IF(pairFlag,otherMask*((ISNUMBER(SEARCH(LEFT(pairText,1),unitVals))+ISNUMBER(SEARCH(RIGHT(pairText,1),unitVals)))>0),SEQUENCE(9)=0),targetCells,_xlws.FILTER("R"&SEQUENCE(9)&"C"&colNum,hitMask,""),targetText,TEXTJOIN(", ",TRUE,targetCells),IF(targetText<>"","Naked pair in column · R"&firstPos&"C"&colNum&" and R"&secondPos&"C"&colNum&" contain only "&LEFT(pairText,1)&"/"&RIGHT(pairText,1)&" · Remove "&LEFT(pairText,1)&" and "&RIGHT(pairText,1)&" from "&targetText&".","")))))))
```

## `SDK_HintNakedPairRow`

```excel
=LAMBDA(board,LET(candMatrix,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(foundHint,scanIndex,IF(foundHint<>"",foundHint,LET(rowNum,QUOTIENT(scanIndex-1,9)+1,colNum,MOD(scanIndex-1,9)+1,pairText,INDEX(candMatrix,rowNum,colNum)&"",unitVals,INDEX(candMatrix,rowNum,0)&"",pairFlag,AND(LEN(pairText)=2,SUM(--(unitVals=pairText))=2),pairPositions,_xlws.FILTER(SEQUENCE(,9),unitVals=pairText,""),firstPos,IF(pairFlag,INDEX(pairPositions,1,1),0),secondPos,IF(pairFlag,INDEX(pairPositions,1,COLUMNS(pairPositions)),0),otherMask,(SEQUENCE(,9)<>firstPos)*(SEQUENCE(,9)<>secondPos),hitMask,IF(pairFlag,otherMask*((ISNUMBER(SEARCH(LEFT(pairText,1),unitVals))+ISNUMBER(SEARCH(RIGHT(pairText,1),unitVals)))>0),SEQUENCE(,9)=0),targetCells,_xlws.FILTER("R"&rowNum&"C"&SEQUENCE(,9),hitMask,""),targetText,TEXTJOIN(", ",TRUE,targetCells),IF(targetText<>"","Naked pair in row · R"&rowNum&"C"&firstPos&" and R"&rowNum&"C"&secondPos&" contain only "&LEFT(pairText,1)&"/"&RIGHT(pairText,1)&" · Remove "&LEFT(pairText,1)&" and "&RIGHT(pairText,1)&" from "&targetText&".","")))))))
```

## `SDK_HintNakedPairRowCandidates`

```excel
=LAMBDA(candMatrix,REDUCE("",SEQUENCE(81),LAMBDA(foundHint,scanIndex,IF(foundHint<>"",foundHint,LET(rowNum,QUOTIENT(scanIndex-1,9)+1,colNum,MOD(scanIndex-1,9)+1,pairText,INDEX(candMatrix,rowNum,colNum)&"",unitVals,INDEX(candMatrix,rowNum,0)&"",pairFlag,AND(LEN(pairText)=2,SUM(--(unitVals=pairText))=2),pairPositions,_xlws.FILTER(SEQUENCE(,9),unitVals=pairText,""),firstPos,IF(pairFlag,INDEX(pairPositions,1,1),0),secondPos,IF(pairFlag,INDEX(pairPositions,1,COLUMNS(pairPositions)),0),otherMask,(SEQUENCE(,9)<>firstPos)*(SEQUENCE(,9)<>secondPos),hitMask,IF(pairFlag,otherMask*((ISNUMBER(SEARCH(LEFT(pairText,1),unitVals))+ISNUMBER(SEARCH(RIGHT(pairText,1),unitVals)))>0),SEQUENCE(,9)=0),targetCells,_xlws.FILTER("R"&rowNum&"C"&SEQUENCE(,9),hitMask,""),targetText,TEXTJOIN(", ",TRUE,targetCells),IF(targetText<>"","Naked pair in row · R"&rowNum&"C"&firstPos&" and R"&rowNum&"C"&secondPos&" contain only "&LEFT(pairText,1)&"/"&RIGHT(pairText,1)&" · Remove "&LEFT(pairText,1)&" and "&RIGHT(pairText,1)&" from "&targetText&".",""))))))
```

## `SDK_HintPointing`

```excel
=LAMBDA(board,SDK_PointingHint(board))
```

## `SDK_HintPointingCol`

```excel
=LAMBDA(board,SDK_PointingColHint(board))
```

## `SDK_HintPointingRow`

```excel
=LAMBDA(board,SDK_PointingRowHint(board))
```

## `SDK_IsSolved`

```excel
=LAMBDA(board,solution,AND(TOCOL(board)=TOCOL(solution)))
```

## `SDK_NextHint`

```excel
=LAMBDA(board,LET(a,SDK_HintNaked(board),b,SDK_HintHiddenRow(board),c,SDK_HintHiddenCol(board),d,SDK_HintHiddenBox(board),IF(a<>"",a,IF(b<>"",b,IF(c<>"",c,IF(d<>"",d,"No naked or hidden single is currently available."))))))
```

## `SDK_PlayerHint`

```excel
=LAMBDA(board,solution,LET(logic,SDK_CoachHintV2(board),IF(logic="No supported logical step is currently available.",SDK_FallbackReveal(board,solution),logic)))
```

## `SDK_PlayerTrace`

```excel
=LAMBDA(board,solution,LET(hintText,SDK_PlayerHint(board,solution),parts,TEXTSPLIT(hintText," · "),partCount,COLUMNS(parts),lineOne,IF(partCount>1,"Technique — "&INDEX(parts,1,1)&" | Pattern — "&INDEX(parts,1,2),"Status — "&INDEX(parts,1,1)),lineTwo,IF(partCount>2,"Action — "&TEXTJOIN(" · ",TRUE,DROP(parts,,2)),"Action — No further supported step."),VSTACK(lineOne,lineTwo)))
```

## `SDK_PointingColHint`

```excel
=LAMBDA(board,LET(candMatrix,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(foundHint,scanIndex,IF(foundHint<>"",foundHint,LET(boxNum,QUOTIENT(scanIndex-1,9)+1,digitNum,MOD(scanIndex-1,9)+1,boxStartRow,1+3*QUOTIENT(boxNum-1,3),boxStartCol,1+3*MOD(boxNum-1,3),boxCandidates,INDEX(candMatrix,SEQUENCE(3,,boxStartRow),SEQUENCE(,3,boxStartCol)),candidateMask,ISNUMBER(SEARCH(digitNum&"",boxCandidates&"")),candidateCount,SUM(--candidateMask),colCoordinates,TOCOL(MAKEARRAY(3,3,LAMBDA(rowOffset,colOffset,boxStartCol+colOffset-1))),candidateCols,_xlws.FILTER(colCoordinates,TOCOL(candidateMask),""),singleCol,IF(candidateCount>1,MIN(candidateCols)=MAX(candidateCols),FALSE),targetCol,IF(singleCol,INDEX(candidateCols,1),1),colCandidates,CHOOSECOLS(candMatrix,targetCol)&"",rowNumbers,SEQUENCE(9),outsideMask,((rowNumbers<boxStartRow)+(rowNumbers>boxStartRow+2))>0,containsMask,ISNUMBER(SEARCH(digitNum&"",colCandidates)),targetCells,_xlws.FILTER("R"&rowNumbers&"C"&targetCol,outsideMask*containsMask,""),targetText,TEXTJOIN(", ",TRUE,targetCells),IF(AND(singleCol,targetText<>""),"Pointing column · Box "&boxNum&" confines digit "&digitNum&" to column "&targetCol&" · Remove "&digitNum&" from "&targetText&".","")))))))
```

## `SDK_PointingHint`

```excel
=LAMBDA(board,LET(rowHint,SDK_PointingRowHint(board),colHint,SDK_PointingColHint(board),IF(rowHint<>"",rowHint,colHint)))
```

## `SDK_PointingRowHint`

```excel
=LAMBDA(board,LET(candMatrix,SDK_CandidateMatrix(board),REDUCE("",SEQUENCE(81),LAMBDA(foundHint,scanIndex,IF(foundHint<>"",foundHint,LET(boxNum,QUOTIENT(scanIndex-1,9)+1,digitNum,MOD(scanIndex-1,9)+1,boxStartRow,1+3*QUOTIENT(boxNum-1,3),boxStartCol,1+3*MOD(boxNum-1,3),boxCandidates,INDEX(candMatrix,SEQUENCE(3,,boxStartRow),SEQUENCE(,3,boxStartCol)),candidateMask,ISNUMBER(SEARCH(digitNum&"",boxCandidates&"")),candidateCount,SUM(--candidateMask),rowCoordinates,TOCOL(MAKEARRAY(3,3,LAMBDA(rowOffset,colOffset,boxStartRow+rowOffset-1))),candidateRows,_xlws.FILTER(rowCoordinates,TOCOL(candidateMask),""),singleRow,IF(candidateCount>1,MIN(candidateRows)=MAX(candidateRows),FALSE),targetRow,IF(singleRow,INDEX(candidateRows,1),1),rowCandidates,CHOOSEROWS(candMatrix,targetRow)&"",columnNumbers,SEQUENCE(,9),outsideMask,((columnNumbers<boxStartCol)+(columnNumbers>boxStartCol+2))>0,containsMask,ISNUMBER(SEARCH(digitNum&"",rowCandidates)),targetCells,_xlws.FILTER("R"&targetRow&"C"&columnNumbers,outsideMask*containsMask,""),targetText,TEXTJOIN(", ",TRUE,targetCells),IF(AND(singleRow,targetText<>""),"Pointing row · Box "&boxNum&" confines digit "&digitNum&" to row "&targetRow&" · Remove "&digitNum&" from "&targetText&".","")))))))
```

## `SDK_RenderDigit`

```excel
=LAMBDA(mode,noteText,legalText,digit,LET(n,noteText&"",l,legalText&"",IF(mode="Auto",IF(ISNUMBER(SEARCH(digit,l)),digit,""),IF(mode="Manual",IF(ISNUMBER(SEARCH(digit,n)),digit,""),IF(n="",IF(ISNUMBER(SEARCH(digit,l)),digit,""),IF(AND(ISNUMBER(SEARCH(digit,n)),ISNUMBER(SEARCH(digit,l))),digit,""))))))
```

## `SDK_SetCell`

Returns a new 9×9 board with one cell replaced while preserving true blanks.

```excel
=LAMBDA(board,r,c,value,MAKEARRAY(9,9,LAMBDA(rr,cc,LET(x,INDEX(board,rr,cc),IF(AND(rr=r,cc=c),value,IF(LEN(x&"")=0,"",x))))))
```

## `SDK_Solve`

Experimental formula-only backtracking solver using minimum remaining candidates and SDK_SolveTry branch recursion.

```excel
=LAMBDA(board,LET(conflicts,SUM(MAKEARRAY(9,9,LAMBDA(r,c,--SDK_Conflict(board,r,c)))),dead,SDK_HasDeadEnd(board),emptyCount,SUM(--(LEN(board&"")=0)),IF(OR(conflicts>0,dead),NA(),IF(emptyCount=0,board,LET(candidates,SDK_CandidateMatrix(board),lengths,IF(LEN(board&"")=0,LEN(candidates),99),minLength,MIN(lengths),position,XMATCH(minLength,TOCOL(lengths)),r,QUOTIENT(position-1,9)+1,c,MOD(position-1,9)+1,options,INDEX(candidates,r,c)&"",SDK_SolveTry(board,r,c,options,1))))))
```

## `SDK_SolveStatus`

```excel
=LAMBDA(board,IF(OR(ROWS(board)<>9,COLUMNS(board)<>9),"SEARCH ERROR",IFERROR(LET(result,SDK_Solve(board),IF(AND(ISNA(result)),"NO SOLUTION","VIABLE")),"SEARCH ERROR")))
```

## `SDK_SolveTry`

Recursive branch iterator used by SDK_Solve; tries candidate digits in order.

```excel
=LAMBDA(board,r,c,options,i,IF(i>LEN(options),NA(),LET(nextBoard,SDK_SetCell(board,r,c,--MID(options,i,1)),result,IFERROR(SDK_Solve(nextBoard),NA()),IF(IFERROR(AND(ROWS(result)=9,COLUMNS(result)=9),FALSE),result,SDK_SolveTry(board,r,c,options,i+1)))))
```

## `SDK_Uniqueness`

```excel
=LAMBDA(board,IFERROR(LET(n,SDK_CountSolutions(board,2),CHOOSE(n+1,"NO SOLUTION","UNIQUE","MULTIPLE SOLUTIONS")),"SEARCH ERROR"))
```
