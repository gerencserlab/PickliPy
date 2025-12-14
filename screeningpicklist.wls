#!/usr/bin/env wolframscript
(* ::Package:: *)

(* ::Input::Initialization:: *)
Print["Picklist generator for screening with acoustic dispensing using Beckman Coulter Echo 650 controlled by Revvity PlateWorks. Written by Akos A. Gerencser."];
Print["Version 6.6  10/23/2025"];
Print["Creates as many destination plates as needed to dispense the entire library with one fixed layout, where library compounds sequentially filled in."];
Print["Creates inventory, process, loop, mergle list and map files."];
Print["The input is picklist.xlsx in the same folder as the script. This file must contain SRC, DST and LIB worsheets and an optional black list worksheet. The format is 384w."];
Print["Supports volume tracking and well switching, multi-compound additions per well using repeated compound label entries, and per compound concentration control."];
Print["Reads survey files if present in the same folder and checks avaliable volumes."];
Print["Script usage:\nwolframscript screeningpicklist.wsl \"full path to xlsx file\""]; 
(*2/13/2025: table merge csv is compatible with values with comma and quotes*)
(*2/15/2025: excel columns (or plate rows) can handle multiple letter codes e.g. AA AAA*)
(*4/4/2025: Abort for missing labels. Resistant to using well names as labels*)
(*5/5/2025: Fixed glitch in error handling and removed warning for extra loop*)
(*6/26/2025: Checking file not found; DST wells Blacklist handling*)
(*7/08/2025: Fixed no blacklist operation, enabled using one label for multiple additions - e.g. adding DMSO for constant vehicle volume*)
(*allows using 0 concentration labels that will be ignored. LIB sheet empty rows are ignored now; first column must have a value to use a row.*)
(*9/10/2025: Fixed bug when using decimal point in concentrations in Library sheets; Added warning for overlapping explicit and library wells; Tracking well volume in each plate; Fixed merge table for multi-compound dispense per well (using multiple compound label entries per well); compound and cocentrations maps can do multiple compounds per well now.*)
(*10/22/2025: Survey files in the same folder are handled. Allows repeated library entries of the same well/compound for concenration control or repeated dispense*)
(*10/23/2025: New addition override to source volume checking*)

src=Null;
dst=Null;
lib=Null;
blk=Null;

fpickist="picklist";
pl=Import[f=$ScriptCommandLine[[2]]];(**)
If[pl===$Failed,Print["Picklist file not found or cannot be imported!"];Print["Press Enter to exit..."];InputString[];Abort[];]
If[Length[pl]==4,{src,dst,lib,blk}=pl[[{1,2,3,4}]],{src,dst,lib}=pl[[{1,2,3}]];blk=Null];
worksheetnames=Import[f,"Sheets"][[{1,2}]];

If[src===Null||dst===Null||lib===Null,Print["Worksheet not found: "<>f];Print["Press Enter to exit..."];InputString[];Abort[];];
src=DeleteCases[src,_?(#[[1]]==""&)];
srcorig=src;
Print[f<>" loaded succesfully."];
SetDirectory[FileNameDrop[f,-1]];

(*remove empty lines from library - first column must have values!*)
lib=Join[lib[[{1}]],Cases[lib[[2;;]],{a_,___}/;StringTrim[ToString[a]]!="",1]];

(*survey files*)
xmlssurvey=FileNames[FileNameDrop[f,-1]<>"\\*_platesurvey.xml"];
vols=<||>;
If[Length[xmlssurvey]>0,
With[{xml=Import[#]},
srcplate=Cases[xml,XMLElement["platesurvey",{___,"barcode"->data_,___},_]:>data,Infinity][[1]];
If[srcplate=="UnknownBarCode",
If[Length[Union[barcodesSRC]]==1,srcplate=barcodesSRC[[1]],
Print["Plate survey file with unknown bardcode has been provided, and this cannot be matched to multiple source barcodes in use."];Print["Press Enter to exit..."];InputString[];Abort[];
]];
printmap=Cases[xml,XMLElement["platesurvey",_,data_]:>data,Infinity];
surveyattrs=Cases[printmap,XMLElement["w",attr_,_]:>attr,Infinity];
vols=Association[vols,<|ReleaseHold[{Hold[srcplate<>"n"->ToExpression["vl"]]}/.surveyattrs]|>];

]&/@xmlssurvey];


towellname[row_, col_] := StringJoin[FromCharacterCode[64 + row], ToString[col]]; 
fw[wellname_] := {ToExpression[StringDrop[wellname, 1]], ToCharacterCode[StringTake[wellname, 1]][[1]] - 64};
tocol[s_] :=26^(Reverse[Range[Length[#]]-1]) . #&@(ToCharacterCode[s]-64); 
torowcol[s_] := StringCases[s, r:LetterCharacter..~~c:DigitCharacter.. :> {tocol[r], 
       ToExpression[c]}][[1]]; 
makeInventory[barcode_, rack_, pos_] := {barcode, "", "", 1, "", "", "", rack, pos, ""}; 
addheaders[tb_] := Join[{Table[i, {i, 0, 12}]}, Join[Table[{FromCharacterCode[i + 64]}, {i, 8}], tb, 2]]; 
SetAttributes[IntegerChop, Listable]; 
IntegerChop[n_] := With[{r = Round[n]}, If[Chop[N[n - r, 100]] == 0, r, n]]; 
IntegerChop[n_, del_] := With[{r = Round[n]}, 
    If[Chop[N[n - r, IntegerLength[r] + Max[Log[10, 1 + Abs[del]] + 3, 100]], del] == 0, r, n]]; 
dilute[{{well_, stock_,___}, final_}] := {well, Round[1000*wellvolume*(final/stock), 2.5]}; 
diluteError[{{well_, stock_,___}, final_}] := 
   If[Abs[1000*wellvolume*(final/stock) - Max[2.5, Round[1000*wellvolume*(final/stock), 2.5]]]/
      (1000*wellvolume*(final/stock)) > 0.1, 
    Print["Warning: Significant rounding error at low volume dispensing! Compound: ", well /. srcName[activesrc], 
      " Stock: ", stock, " Final: ", final, " Desired dispense volume: ", 1000*wellvolume*(final/stock), 
      " Possible dispense volume: ", Max[2.5, Round[1000*wellvolume*(final/stock), 2.5]]]; warnings = True]; 
torowcol[s_] := StringCases[s, r:LetterCharacter~~c:DigitCharacter.. :> {ToCharacterCode[r][[1]] - 64, 
       ToExpression[c]}][[1]]; 
reordersrc[list_] := With[{o = OrderingBy[list, N[EuclideanDistance[{1, 1}, torowcol[#1[[1,2,1]]]]] & ]}, 
    list[[FindShortestTour[torowcol /@ list[[All,1,2,1]], First[o], Last[o], 
       DistanceFunction -> EuclideanDistance, Method -> "Greedy"][[2]]]]]; 
reorderdst[list_] := With[{o = OrderingBy[list, N[EuclideanDistance[lastpos, torowcol[#1[[1]]]]] & ]}, 
    lastpos = torowcol[list[[Last[o],1]]]; list[[FindShortestTour[torowcol /@ list[[All,1]], First[o], Last[o], 
        DistanceFunction -> EuclideanDistance, Method -> "Greedy"][[2]]]]]; 
chooseSrc[{dstwell_, srcwells_, vol_}] := 
   Module[{}, s = Select[srcwells, volumes[StringJoin[activesrc, "_", #1]] >= minvol + vol/1000 & , 1]; 
     If[Length[s] == 0, Throw[StringJoin["Error: Out of Volume for ", srcwells[[1]] /. srcName[activesrc]]]]; 
     volumes[StringJoin[activesrc, "_", s[[1]]]] -= vol/1000; {dstwell, s[[1]], vol}];
addcompound[srcline_]:=Module[{notyet},
notyet=True;
If[libindex<Length[lib]-1,
Append[srcline/.s_String:>StringReplace[s,"@"~~x:LetterCharacter..:>lib[[
If[notyet,notyet=False;++libindex,libindex]+1
,tocol[x]]]]/.StringExpression[st_]:>st,libindex],
{}]
];

header=Map[#[[1]]->#[[2]]&,dst[[1;;12,{1,2}]]];
InventorySRC=<|header|>["Inventory_SRC:"]/.Missing[__]->"Inventory";
InventoryDST=<|header|>["Inventory_DST:"]/.Missing[__]->"Inventory";
SRCRack=IntegerChop[<|header|>["Inventory_SRC_Rack#:"]/.Missing[__]->1];
DSTRack=IntegerChop[<|header|>["Inventory_DST_Rack#:"]/.Missing[__]->2];
ProcessSRC=<|header|>["Process_SRC:"]/.Missing[__]->"Process_SRC";
ProcessDST=<|header|>["Process_DST:"]/.Missing[__]->"Process_DST";
LoopSRC=<|header|>["Loop Counts_SRC:"]/.Missing[__]->"Loop_SRC";
LoopDST=<|header|>["Loop Counts_DST:"]/.Missing[__]->"Loop_DST";
wellvolume=ToExpression[<|header|>["Well volume (ul):"]/.Missing[__]->dst[[1,2]]];
If[!NumberQ[wellvolume],Print["Error: missing definition of assay (DST) well volume 'Well volume (ul):'"];Print["Press Enter to exit..."];InputString[];Abort[];];
minvol=src[[1,7]];
If[!NumberQ[minvol],Print["Error: missing minimal Echo Source well volume in 'src' wotksheet G1!"];Print["Press Enter to exit..."];InputString[];Abort[];];

p0=Count[dst[[All,1]],"Plate Map:"];
p1=SequencePosition[dst[[All,1]],Shortest[{"Plate Map:",(z__/;z!=""),"Labels"}]];
parameters=Map[#[[1]]->#[[2]]&,(dst[[#[[1]]+17;;#[[2]]-1,1;;2]]&/@p1),{2}];
additions=(<|#|>["Picklist Name:"]&/@parameters);
barcodesDST=(<|#|>["Barcode_DST:"]&/@parameters)/.Missing[__]->"PlateDST";
allBbarcodesDST=StringSplit[barcodesDST[[1]],","];
If[allBbarcodesDST=={},Print["Barcode_DST definition is missing. Cannot use * for all barcodes."];Print["Press Enter to exit..."];InputString[];Abort[];];
tables=p1[[All,1]];
activesrc="";
lastpos={1,1};
p2={#,#+Position[Append[dst[[#;;,1]],""],""|"Plate Map:"][[1,1]]-2}&/@Flatten[Position[dst[[All,1]],"Labels"]];(*fixed in 4.3 added ,___, and &&z!="Plate Map:"*)

(*labellists will contain compound labels and concentrations for each addition*) 
labellists00=dst[[#[[1]]+1;;#[[2]],1;;3]]&/@p2;
(*Zero concentration labels are deleted, so cause no problems*)
labellists00=DeleteCases[labellists00,{_,_,0.},2];

a=dst[[1+tables[[1]];;16+tables[[1]],2;;25]];
a=a/.x_Real:>IntegerChop[x];
a=a/.x_ /; NumberQ[x]:>ToString[x];

b0={towellname[##],a[[##]]}&@@@Position[a,x_String/;x!=""];


(*blacklist handling*)
If[!blk===Null,
p1=SequencePosition[blk[[All,1]],Shortest[{"Groups:",(z__/;z!=""),"Barcode_DST"}]];
blktables=p1[[All,1]];
p2={#,#+Position[Append[blk[[#;;,1]],""],""|"Groups:"][[1,1]]-2}&/@Flatten[Position[blk[[All,1]],"Barcode_DST"]];
labellists0blk=blk[[#[[1]]+1;;#[[2]],1;;2]]&/@p2;
labellistsblk={StringSplit[#[[1]],{",",", "," , "}],StringSplit[#[[2]],{",",", "," , "}]}&/@labellists0blk[[1]];
labellistsblk=Flatten[Outer[Rule,#[[1]],#[[2]]]&/@labellistsblk,2];
labellistsblkdict=<|#[[1,1]]->#[[All,2]]&/@GatherBy[labellistsblk,#[[1]]&]|>;

a=blk[[1+blktables[[1]];;16+blktables[[1]],2;;25]];
a=a/.x_Real:>IntegerChop[x];
a=a/.x_ /; NumberQ[x]:>ToString[x];
b={towellname[##],a[[##]]}&@@@Position[a,x_String/;x!=""];
blkgroupsdict=<|{#[[1,2]]->#[[All,1]]}&/@GatherBy[b,#[[2]]&]|>;
blkwellsdict=<|{#[[1]]->#[[2]]}&/@b|>;
labeldict=<|#[[1]]->#[[2]]&/@labellists00[[1]]|>;
labeldictrev=<|#[[1,2]]->#[[All,1]]&/@GatherBy[labellists00[[1]],#[[2]]&]|>;
];


oneaddition[n_]:=Module[{},
Print["****** ",additions[[1]]<>".csv"," ******"];
loopsrc=Append[loopsrc,barcodesSRC[[n]]];
loopdst=Append[loopdst,barcodeDST];
If[MissingQ[conditionlabels[barcodeDST]],conditionlabels[barcodeDST]=Table["",{16},{24}]];
Print["Dispensing from: "<>barcodesSRC[[n]]];
Print["Dispensing into: "<>barcodeDST];
Off[StringMatchQ::nonopt];
activesrc=barcodesSRC[[n]];
diluteError[#[[{2,3}]]]&/@labellists[[n]];
l=#[[1]]->dilute[#[[{2,3}]]]&/@labellists[[n]];
l=(#[[1,1]]->Flatten[#[[All,2]],1])&/@GatherBy[l,#[[1]]&];(*mod for multi compound per label*)
labelsdecoded = c/.{x_String,y_}:>{x,y/.l}; (*250402 fix for using well labels as labels*)
(*cl=Complement[labelsdecoded,c];*)cl=Flatten[Complement[labelsdecoded,c]/.{x_,y_}:>({x,#}&/@Partition[y,2]),1];(*mod for multi compound per label*)
d=Gather[Flatten[#,1]&/@cl,#1[[2]]==#2[[2]]&];
lname=#[[1]]->{barcodeDST,#[[2,1,1]]/.srcName[barcodesSRC[[n]]],#[[3]]}&/@labellists[[n]];

(*megrge tabel fix 250910 - well labels are unique, compound labels not for multi compound*)
l=#[[1]]->{{#[[2,1,1]]},#[[3]]}&/@labellists[[n]]; (*modified to yield concentrations*)
l=(#[[1,1]]->Flatten[#[[All,2]],1])&/@GatherBy[l,#[[1]]&];(*must match the above!*)
labelsdecoded = c/.{x_String,y_}:>{x,y/.l}; (*must match the above!*)
clname=Flatten[Complement[labelsdecoded,c]/.{x_,y_}:>({x,#}&/@Partition[y,2]),1]; (*must match the above!*)
clname[[All,2,1]]=clname[[All,2,1,1]]/.srcName[barcodesSRC[[n]]];(*megrge tabel fix 250910*)
clname[[All,2]]=Join[{barcodeDST},#]&/@clname[[All,2]];(*megrge tabel fix 250910*)

tablemergeassociations=Join[tablemergeassociations,clname];
dstlabels=conditionlabels[barcodeDST];
(dstlabels[[##&@@torowcol[#[[1]]]]]=If[(a=dstlabels[[##&@@torowcol[#[[1]]]]])==="",{#[[2,2]],#[[2,3]]},{a[[1]]<>"+"<>#[[2,2]],Flatten[{a[[2]],#[[2,3]]}]}])&/@clname;
conditionlabels[barcodeDST]=dstlabels;
(*Abort[];Abort[];Abort[];*)

lastpos={1,1};
optim=reorderdst/@reordersrc[d];
optim=chooseSrc/@Flatten[optim,1];
optim=GatherBy[optim,#[[2]]&];

v={#[[1,2]]/.srcName[barcodesSRC[[n]]],barcodesSRC[[n]]<>"_"<>#[[1,2]],Total@#[[All,3]]/1000//N}&/@optim;
(*
Print["Volumes used:"];
Print[StringRiffle[#,"   "]]&/@v;
*)
totaluse=Append[totaluse,v];

(*Print[ListLinePlot[r=torowcol/@#[[All,1]],PlotMarkers->{Automatic, Small}]&/@optim];*)

csv=Join[csv,Flatten[{barcodesSRC[[n]],#[[1,2]],barcodeDST,#[[All,{3,1}]]}]&/@optim];
];

fixOneWellBlackList[badwelllabel_]:=Module[{},
(*Print["Checking: ",badwelllabel];*)
If[MissingQ[blkwellsdict[badwelllabel]],Return[Null]]; (*Exclude controls or fixed wells from groups in Groups Map!*)
badwell=Cases[b,_?(#[[1]]==badwelllabel&)];
If[Length[badwell]==0,Print["Blacklisting: ",badwelllabel, " - it is not on the dispense list any more, skipping..."];Return[Null],badwell=badwell[[1]]];
(*Print["Blacklisting: ",badwell[[1]]];*)
removecompound=labeldictrev[labeldict[badwell[[2]]]];
removewells=Cases[b,_?(MemberQ[removecompound,#[[2]]]&)][[All,1]];
replacementwell=freewells[blkwellsdict[badwelllabel]];
If[Length[replacementwell]>0,
freewells[blkwellsdict[badwelllabel]]=DeleteElements[freewells[blkwellsdict[badwelllabel]],replacementwell[[{1}]]];
replacementwell=replacementwell[[1]];
fixedwell=badwell;
fixedwell[[1]]=replacementwell;
b=b/.badwell->fixedwell;
Print["Switching ",badwell[[1]]," to ",replacementwell];
,
sparewells=Complement[Cases[b,_?(MemberQ[removewells,#[[1]]]&)][[All,1]],{badwelllabel}];
sparewells=Complement[sparewells,blacklistedwells]; (*blacklisted wells are not added to the available pool, these this will be skipped later on*)If[Length[sparewells]>0,(freewells[blkwellsdict[#]]=Append[freewells[blkwellsdict[#]],#])&/@sparewells];
b=DeleteCases[b,_?(MemberQ[removewells,#[[1]]]&)];
Print["Removing all instances of ",removecompound, " wells: ",StringRiffle[removewells,", "]];
]
];

looplibrary[]:=Module[{},
dstcount++;
If[Length[allBbarcodesDST]<dstcount,Print["Error: please provide more destination bar codes in the DST worksheet Barcode_DST: field"];Print["Press Enter to exit..."];InputString[];Abort[];];
barcodeDST=allBbarcodesDST[[dstcount]]; (*New DST plate*)



b=b0;

(*blacklist handling*)
If[!blk===Null,
blacklistedwells=labellistsblkdict[barcodeDST];
clnameblk={};

If[!MissingQ[blacklistedwells],
Print["****** Preparing Blacklist for: ",barcodeDST," ******"];
freewells=<|Replace[Normal[blkgroupsdict],x_List:>Complement[x,b[[All,1]],blacklistedwells],2]|>; (*this is useless for the last plate, at this point is unknown where will the library end....*)
fixOneWellBlackList/@blacklistedwells;
clnameblk={#,{barcodeDST,"rejected",0}}&/@blacklistedwells;
tablemergeassociationsblk=Join[tablemergeassociationsblk,clnameblk];
Print["Adding remaining free wells as ",labellists00[[1,-1,1]],": ",freewells];
b=Join[b,{#,labellists00[[1,-1,1]]}&/@Flatten@Values@freewells];
];
];

(*moved here from the top*)
c=Flatten[b/.{well_String,treat_String}:>({well,#}&/@StringSplit[treat,{"+",",",";","\n"}]),1];
u=Union[DeleteElements[Flatten[b[[All,2]]],{""}]];
labellists0=DeleteCases[labellists00,{x_/;!MemberQ[u,x],_,_},2];
compoundsused=labellists0[[1,All,2]];
missinglabels = Complement[u,labellists0[[1,All,1]]];
If[(missinglabels!={})||MemberQ[Union[additions,barcodesDST],Missing[__]],Print["Some definitions are missing. Each plate map must start with 'Plate Map:' and followed by lines defining:\nPicklist Name:\nBarcode_SRC:\nBarcode_DST:\nLabels"];
Print["All labels used in the Plate Map must be defined in the Labels list. The Labels list may contain unused entries."];
Print["Missing labels:",missinglabels];
Print["Press Enter to exit..."];InputString[];Abort[]];



wildcardrow=Cases[src[[2;;]],_?(#[[2]]=="*"&),1];
If[Length[wildcardrow]>=1,
explicitcompounds=Cases[src[[2;;]],_?(!StringMatchQ[#[[1]],"@"~~__]&)];
blankcompounds=DeleteDuplicates[DeleteElements[compoundsused,explicitcompounds[[All,2]]]];
allblankrows=wildcardrow[[1]]/."*"->#&/@blankcompounds;
src2=Join[explicitcompounds,allblankrows];
src2=addcompound/@src2,
src2=addcompound/@src[[2;;]];
];



src2=DeleteCases[src2,{}];
src2[[All,6]]=src2[[All,6]]/."*"->DeleteCases[Union[src2[[All,6]]],"*"][[1]];

srcorig=src2;
srccol=5;
If[Length[vols]>0,src2[[2;;,srccol]]=(vols[#[[1]]<>#[[2]]]&/@src2[[2;;,{6,1}]])/.Missing[___]->0;Print["Using survey volumes: ",src2[[2;;,{1,6,4,5}]]//TableForm]];

newlyaddedwells=Position[src2[[2;;,7]],"new"]//Flatten;
If[Length[newlyaddedwells]>0,src2[[1+newlyaddedwells,srccol]]=srcorig[[1+newlyaddedwells,srccol]]];

srcr=(#[[1,6]]->((#[[1,1]]->{#[[All,2,1]],#[[1,2,2]],#[[1,2,3]]})&/@GatherBy[#[[2]]->Append[#[[{1,3}]],Last[#]]&/@#,#[[1]]&]))&/@GatherBy[src2,#[[6]]&];
barcodesSRC=Tally[src2[[All,6]]][[All,1]];(*Tally keeps original order, Union does not*)
srcName=<|#[[1,6]]->((#[[1]]->#[[4]])&/@#)&/@GatherBy[src2,#[[6]]&]|>;
volumes=Association[<|(#[[6]]<>"_"<>#[[1]]->#[[5]])&/@src2|>,volumes]; (*250910 fix for multi-plate volume tracking*)

duplicateSources=Gather[src2,#1[[{1,6}]]==#2[[{1,6}]]&&#1[[4]]!=#2[[4]]&];
If[Max[Length/@duplicateSources]>1,Print["Error in SRC sheet - explicit compound (control) wells overlap with library wells, or library has duplicate well definitions for a barcode!"];
Print[Cases[duplicateSources,_?(Length[#]>1&)][[All,All,{6,1,4}]]];
InputString[];Abort[];];

If[lastbarcode!=""&&lastbarcode!=First[barcodesSRC],loopdst=Append[loopdst,barcodeDST];
(*
warnings=True;Print["Warning: extra DST loop added to trigger plate switching in Plate:Works, because SRC plate completed together with completing DST plate."]
*)
];

lastbarcode=Last[barcodesSRC];
(*note: the code below supports only a single platemap*)
labellists=Table[labellists0[[1]],Length[barcodesSRC]];

(*Apply indirect concetration references made in the DST worksheet referring to the compounds concentration in a LIB table comlumn. This require the 3rd value in decoded srcr to be the library row (without header)*)
compounds=Flatten[barcodesSRC/.srcr,1][[All,1]];
labellists=labellists/.{label_,compound_,concentration_String}/;MemberQ[compounds,compound]:>{label,compound,lib[[(compound/.Flatten[barcodesSRC/.srcr,1])[[3]]+1,tocol[StringReplace[concentration,"@"~~x:LetterCharacter..:>x]]]]};
(*decode SRC positions and stock concentrations*)
labellists[[All,All,2]]=MapThread[ReplaceAll,{labellists[[All,All,2]],barcodesSRC/.srcr}];
labellists=Cases[#,{_,{__},_}]&/@labellists;


Do[oneaddition[i],{i,Length[barcodesSRC]}];

If[Length[clnameblk]>0,
dstlabels=conditionlabels[barcodeDST];
(dstlabels[[##&@@torowcol[#[[1]]]]]={"rejected",0})&/@clnameblk;
conditionlabels[barcodeDST]=dstlabels;
];


];

lastbarcode="";
conditionlabels=<|{}|>;
tablemergeassociations={};
tablemergeassociationsblk={};
loopsrc={};
loopdst={};
dstcount=0;
csv={};
libindex=0;
additions0="";
totaluse={};
warnings=False;
volumes =<||>;
Print[
Catch[

While[libindex<Length[lib]-1,looplibrary[]];

Print["Saving: ",Export[additions[[1]]<>".csv",StringRiffle[csv,"\n",","],"Text"]];

Print["****** Summary: Total Usage ******"];
Print[#[[1,1]],"\t\t",#[[1,2]],"\t",Total[#[[All,3]]],"\tremains (ul): ",volumes[#[[1,2]]]]&/@GatherBy[Flatten[totaluse,1],#[[2]]&];

Print["Dispensed compounds (with controls)= ",Length[Flatten[totaluse,1]]];

Print["***** Plateworks Tables ******"];

inventory={
MapIndexed[makeInventory[#1,SRCRack,#2[[1]]]&,Tally[loopsrc][[All,1]]],
MapIndexed[makeInventory[#1,DSTRack,#2[[1]]]&,Tally[loopdst][[All,1]]]};
If[InventorySRC==InventoryDST,inventory={Join[##&@@inventory]}];(*fix 4.3 the Join did not work properly, one device inventory wasn't possible before*)
Print["***** Inventory = "<>InventorySRC<>".csv"];
inventorytable=StringRiffle[inventory[[1]],"\n",","];
Print[inventorytable];
Export[InventorySRC<>".csv",inventorytable,"Text"];

If[InventorySRC!=InventoryDST,
Print["***** Inventory = "<>InventoryDST<>".csv"];
inventorytable=StringRiffle[inventory[[2]],"\n",","];
Print[inventorytable];
Export[InventoryDST<>".csv",inventorytable,"Text"];
];


Print["***** Process_SRC = "<>ProcessSRC<>".txt"];
processtable=StringRiffle[Tally[loopsrc][[All,1]],"\n"];
Print[processtable];
Export[ProcessSRC<>".txt",processtable,"Text"];
Print["***** Process_DST = "<>ProcessDST<>".txt"];
processtable=StringRiffle[Tally[loopdst][[All,1]],"\n"];
Print[processtable];
Export[ProcessDST<>".txt",processtable,"Text"];

(*
Print["***** Loop_SRC = "<>LoopSRC<>".csv"];
looptable=StringRiffle[Join[{{"Barcode","V"}},Tally[loopsrc]],"\n",","];
Print[looptable];
Export[LoopSRC<>".csv",looptable,"Text"];

Print["***** Loop_DST = "<>LoopDST<>".csv"];
looptable=StringRiffle[Join[{{"Barcode","V"}},Tally[loopdst]],"\n",","];
Print[looptable];
Export[LoopDST<>".csv",looptable,"Text"];
*)

Print["***** Table Merge Associations file = "<>"table merge"<>".txt"];
(*quote only expressions with commas for less clutter, do it the same way as Excel saves CSV, the Wolfram CSV save is different, quotes everything*)
tablemergeassociations=Join[tablemergeassociations,tablemergeassociationsblk];
mergetable=StringRiffle[(Flatten/@tablemergeassociations)/.x_String/;StringMatchQ[x,___~~","|"\""~~___]:>"\""<>StringReplace[x,"\""->"\"\""]<>"\"","\n",","]; 
Print[Length[tablemergeassociations]," merge associatinons saved."];
Export["table merge.txt",mergetable,"Text"];

Print["***** Destination plate map file = "<>"plate map compounds"<>".txt"];
table=StringRiffle[Flatten[Join[{#[[1]]},#[[2]]/.{s_String,c_}:>s]&/@Normal[conditionlabels],1],"\n","\t"];
(*Print[table//TableForm];*)
Export["plate map compounds.txt",table,"Text"];

Print["***** Destination plate map file = "<>"plate map concentrations"<>".txt"];
table=StringRiffle[Flatten[Join[{#[[1]]},#[[2]]/.{s_String,c_}:>c]&/@Normal[conditionlabels],1],"\n","\t"];
(*Print[table//TableForm];*)
Export["plate map concentrations.txt",table,"Text"];
Print["CellPlateCount=",Length[Tally[loopdst]]];
Print["LibraryPlateCount=",Length[Tally[loopsrc]]];
Print["TotalDispenseEvents=",Length[loopdst]];

If[warnings,"Finished with warnings. Revise source file if needed!","Picklists passed all checks. Ready."]
]]


Print["Press Enter to exit..."];
InputString[];



(* ::Input:: *)
(**)
