#!/usr/bin/env wolframscript
(* ::Package:: *)

(* ::Input:: *)
(*(*v4.9 4/4/2025: Resistant to using well names as labels*)*)
(*(*v4.10 6/18/2025: Checking SRC for missing labels*)*)
(*(*v4.11 6/26/2025: Checking file not found*)*)
(*(*v4.12 8/14/2025: Optional Layout map added for additional information for condition labels*)*)
(*(*v4.13 9/14/2025: Waits after abort for all errors,table merge, and per barcode conditions list (only one plating conditions is allowed for now)*)*)
(*(*v4.14 9/19/2025: Merge table and Per plate concentrations map both using labels and compound name_concentration*)*)
(*(*v4.15 9/21/2025: Quoted commas in the merge table csv*)*)
(*(*v4.16 9/28/2025: Source survey volumes are used + fix to error checking*)*)
(*(*v4.17 10/23/2025: New addition override to source volume checking*)*)


(* ::Input::Initialization:: *)
Print["Picklist generator for acoustic dispensing using Beckman Coulter Echo 650 controlled by Revvity PlateWorks. Written by Akos A. Gerencser."];
Print["Version 4.17  10/23/2025"];
Print["Supporting multiple source and destination plates, volume checking and distributing larger total volume additions into multiple wells."];
Print["Supporting dose-response, use 'Concentrations Map:' label below the plate map to provide concentrations. For dose response, only one compound per well is supported in one table, dispense multiple compounds using a new plate map with the same picklist name."];
Print["Creates inventory and process files. Source plates go on Rack 1 from bottom up, destination plates go on Rack 2 from bottom up. Barcodes will come alphabetically."];
Print["The input is picklist.xlsx in the same folder as the script. The format is 384w."];
Print["Script usage:\nwolframscript assaypicklist.wsl \"full path to xlsx file\""]; 

src=Null;
dst=Null;

fpickist="picklist";
pl=Import[f=$ScriptCommandLine[[2]]];
If[pl===$Failed,Print["Picklist file not found or cannot be imported!"];Print["Press Enter to exit..."];InputString[];Abort[];]
{src,dst}=pl[[{1,2}]];
worksheetnames=Import[f,"Sheets"][[{1,2}]];

If[src===Null||dst===Null,Print["Worksheet not found: "<>f];Print["Press Enter to exit..."];InputString[];Abort[];];
SetDirectory[FileNameDrop[f,-1]];
srcorig=src;
(*source well labels must be unique*)
sourcewells=Cases[StringJoin/@src[[2;;,{6,1}]],s_/;s!=""];
duplicatewells=Cases[Tally[sourcewells],_?(#[[2]]>1&)];
If[Length[duplicatewells]>0,Print["Error: duplicate well assignement in 'src' worksheet within the same source plate: "<>StringRiffle[duplicatewells[[All,1]],", "]];Print["Press Enter to exit..."];InputString[];Abort[];]
(*duplicate compounds must have the same concentration*)
sourcewells=Cases[src[[2;;,{2,3,6}]],{s_,conc_,plate_}/;s!="":>{plate<>s,conc}];
If[Max[Length[Tally[#]]&/@GatherBy[sourcewells,#[[1]]&]]>1,Print["Error: duplicate compound labels in 'src' worksheet must have the same concentrations within each source plate!"];Print["Press Enter to exit..."];InputString[];Abort[];]
Print[f<>" loaded succesfully."]

towellname[row_,col_]:=FromCharacterCode[64+row]<>ToString[col];
fw[wellname_]:={ToExpression[StringDrop[wellname,1]],ToCharacterCode[StringTake[wellname,1]][[1]]-64}
torowcol[s_]:=StringCases[s,r:LetterCharacter~~c:DigitCharacter..:>{ToCharacterCode[r][[1]]-64,ToExpression[c]}][[1]];
makeInventory[barcode_,rack_,pos_]:={barcode,"","",1,"","","",rack,pos,""};
addheaders[tb_]:=Join[{Table[i,{i,0,12}]},Join[Table[{FromCharacterCode[i+64]},{i,8}],tb,2]];
SetAttributes[IntegerChop,Listable];
IntegerChop[n_]:=With[{r=Round[n]},If[Chop[N[n-r,100]]==0,r,n]];
IntegerChop[n_,del_]:=With[{r=Round[n]},If[Chop[N[n-r,IntegerLength[r]+Max[Log[10,1+Abs[del]]+3,100]],del]==0,r,n]];

header=Map[#[[1]]->#[[2]]&,dst[[1;;8,{1,2}]]];
InventorySRC=<|header|>["Inventory_SRC:"]/.Missing[__]->"Inventory";
InventoryDST=<|header|>["Inventory_DST:"]/.Missing[__]->"Inventory";
SRCRack=IntegerChop[<|header|>["Inventory_SRC_Rack#:"]/.Missing[__]->1];
DSTRack=IntegerChop[<|header|>["Inventory_DST_Rack#:"]/.Missing[__]->2];
ProcessSRC=<|header|>["Process_SRC:"]/.Missing[__]->"Process_SRC";
ProcessDST=<|header|>["Process_DST:"]/.Missing[__]->"Process_DST";
wellvolumeDefault=ToExpression[<|header|>["Well volume (ul):"]/.Missing[__]->dst[[1,2]]];
If[!NumberQ[wellvolumeDefault],Print["Error: missing definition of assay (DST) well volume 'Well volume (ul):'"];Print["Press Enter to exit..."];InputString[];Abort[]];
wellvolume=wellvolumeDefault;
minvol=src[[1,7]];
If[!NumberQ[minvol],Print["Error: missing minimal Echo Source well volume in 'src' wotksheet G1!"];Print["Press Enter to exit..."];InputString[];Abort[]];

p1Layout=Position[dst[[All,1]],"Layout:"];
p0=Count[dst[[All,1]],"Plate Map:"];
p1={#,#+Position[Append[dst[[#;;,1]],""],""|"Labels"][[1,1]]-1}&/@Flatten[Position[dst[[All,1]],"Plate Map:"]];
parameters=Map[#[[1]]->#[[2]]&,(dst[[#[[1]]+1;;#[[2]]-1,1;;2]]&/@p1),{2}];
additions=(<|#|>["Picklist Name:"]&/@parameters);
barcodesSRC=(<|#|>["Barcode_SRC:"]&/@parameters)/.Missing[__]->"PlateSRC";
barcodesDST=(<|#|>["Barcode_DST:"]&/@parameters)/.Missing[__]->"PlateDST";
wellvolumes=(<|#|>["Well volume (ul):"]&/@parameters)/.Missing[__]->wellvolumeDefault;
barcodesDST=StringSplit[barcodesDST,","];
allBbarcodesDST=DeleteCases[Union[Flatten[barcodesDST]],"*"];
If[allBbarcodesDST=={},Print["Barcode_DST definition is missing. Cannot use * for all barcodes."];Print["Press Enter to exit..."];InputString[];Abort[];];

xmlssurvey=FileNames[FileNameDrop[f,-1]<>"\\*_platesurvey.xml"];

srcorig=src;
srccol=5;
If[Length[xmlssurvey]>0,
With[{xml=Import[#]},
srcplate=Cases[xml,XMLElement["platesurvey",{___,"barcode"->data_,___},_]:>data,Infinity][[1]];
If[srcplate=="UnknownBarCode",
If[Length[Union[barcodesSRC]]==1,srcplate=barcodesSRC[[1]],
Print["Plate survey file with unknown bardcode has been provided, and this cannot be matched to multiple source barcodes in use."];Print["Press Enter to exit..."];InputString[];Abort[];
]];
printmap=Cases[xml,XMLElement["platesurvey",_,data_]:>data,Infinity];
surveyattrs=Cases[printmap,XMLElement["w",attr_,_]:>attr,Infinity];
vols=<|ReleaseHold[{Hold[srcplate<>"n"->ToExpression["vl"]]}/.surveyattrs]|>;
If[!(And@@(NumberQ[vols[#[[1]]<>#[[2]]]]&/@src[[2;;,{6,1}]])),Print["Using survey file: ",xmlssurvey," but not finding values for one or more wells: ",vols[#[[1]]<>#[[2]]]&/@src[[2;;,{6,1}]]," Using zero for these"];];
Print["Volumes updated from survey file: ",
src[[2;;,srccol]]=(vols[#[[1]]<>#[[2]]]&/@src[[2;;,{6,1}]])/.Missing[___]->0];
]&/@xmlssurvey];

newlyaddedwells=Position[src[[2;;,7]],"new"]//Flatten;
If[Length[newlyaddedwells]>0,src[[1+newlyaddedwells,srccol]]=srcorig[[1+newlyaddedwells,srccol]]];

tables=p1[[All,1]];
p1c=SequencePosition[dst[[All,1]],Shortest[{"Plate Map:",(z__/;z!=""),"Concentrations Map:"|"Concentrations  Map:"}]]; (*correcting for a typo during intital design*)
ctables=<|First[#]->Last[#]&/@p1c|>;
activesrc="";
lastpos={1,1};
p2={#,#+Position[Append[dst[[#;;,1]],""],""|"Plate Map:"][[1,1]]-2}&/@Flatten[Position[dst[[All,1]],"Labels"]];(*fixed in 4.3 added ,___, and &&z!="Plate Map:"*)

(*labellists will contain compound labels and concentrations for each addition*) 
labellists=dst[[#[[1]]+1;;#[[2]],1;;3]]&/@p2;
If[!(p0==Length[p1]==Length[p2]==Length[tables]==Length[additions]==Length[labellists]==Length[barcodesSRC]==Length[barcodesDST])||MemberQ[Union[additions,barcodesSRC,barcodesDST],Missing[__]],Print["Some definitions are missing. Each plate map must start with 'Plate Map:' and followed by lines defining:\nPicklist Name:\nBarcode_SRC:\nBarcode_DST:\nLabels"];Print["Press Enter to exit..."];InputString[];Abort[];];

(*Rules for compound locations for each unique source plate*)
srcr=(#[[1,6]]->((#[[1,1]]->{#[[All,2,1]],#[[1,2,2]]})&/@GatherBy[#[[2]]->#[[{1,3}]]&/@#,#[[1]]&]))&/@GatherBy[src[[2;;]],#[[6]]&];

(*Each addition is using one (a unique or the same) source plate, this is applied here, but source plate names are not carried forward in the labellists*)
If[!MemberQ[srcr[[All,1]],barcodesSRC[[1]]],Print["The ",barcodesSRC[[1]]," barcode does not exist in the SRC worksheet Col F!"];Print["Press Enter to exit..."];InputString[];Abort[];];

missingsrc=Complement[Union[labellists[[All,All,2]]//Flatten],Union[(barcodesSRC/.srcr)[[All,All,1]]//Flatten]];
If[missingsrc!={},Print["Missing definitons on SRC: ",StringRiffle[missingsrc,", "]];warnings=True;Print["Press Enter to exit..."];InputString[];Abort[];];

compoundlabels=Map[#[[1]]->#[[2]]<>"_"<>ToString[IntegerChop[#[[3]]]]&,labellists,{2}];
compoundlabels=((#[[1,1]]->StringRiffle[Flatten[#[[All,2]],1],"+"])&/@GatherBy[#,#[[1]]&])&/@compoundlabels; (*for multiple compounds per label*)
labellists[[All,All,2]]=MapThread[ReplaceAll,{labellists[[All,All,2]],barcodesSRC/.srcr}];

srcName=<|#[[1,6]]->((#[[1]]->#[[2]])&/@#)&/@GatherBy[src[[2;;]],#[[6]]&]|>;
volumes=<|(#[[6]]<>"_"<>#[[1]]->#[[5]])&/@src[[2;;]]|>;


dilute[{{well_,stock_},final_}]:={well,Round[1000*wellvolume*final/stock,2.5]};
diluteError[{{well_,stock_},final_}]:=Module[{},
If[Abs[1000*wellvolume*final/stock-dilute[{{well,stock},final}][[2]]]/(1000*wellvolume*final/stock)>0.1,Print["Warning: Significant rounding error at low volume dispensing! Compound: ",well/.srcName[activesrc], " Stock: " ,stock," Final: ",final," Desired dispense volume: ", 1000*wellvolume*final/stock," Possible dispense volume: ",dilute[{{well,stock},final}][[2]]];warnings=True];
If[dilute[{{well,stock},final}][[2]]==0,Print["Error: Zero-volume dispensing! Compound: ",well/.srcName[activesrc], " Stock: " ,stock," Final: ",final," Desired dispense volume: ", 1000*wellvolume*final/stock," Possible dispense volume: ",dilute[{{well,stock},final}][[2]]]; Print["Press Enter to exit..."];InputString[];Abort[];
];
];
torowcol[s_]:=StringCases[s,r:LetterCharacter~~c:DigitCharacter..:>{ToCharacterCode[r][[1]]-64,ToExpression[c]}][[1]];
reordersrc[list_]:=With[{o=OrderingBy[list,(EuclideanDistance[{1,1},torowcol[#[[1,2,1]]]]//N)&]},
list[[FindShortestTour[torowcol/@list[[All,1,2,1]],First[o],Last[o],DistanceFunction->EuclideanDistance,Method->"Greedy"][[2]]]]];
reorderdst[list_]:=With[{o=OrderingBy[list,(EuclideanDistance[lastpos,torowcol[#[[1]]]]//N)&]},lastpos=torowcol@list[[Last[o],1]];list[[FindShortestTour[torowcol/@list[[All,1]],First[o],Last[o],DistanceFunction->EuclideanDistance,Method->"Greedy"][[2]]]]];
chooseSrc[{dstwell_,srcwells_,vol_}]:=Module[{},
s=Select[srcwells,volumes[activesrc<>"_"<>#]>=minvol+vol/1000&,1];
If[Length[s]==0,Throw["Error: Out of Volume for "<>(srcwells[[1]]/.srcName[activesrc])]];
volumes[activesrc<>"_"<>s[[1]]]-=vol/1000;
{dstwell,s[[1]],vol}
];
conditionlabels=Table["",{16},{24}];
If[Length[p1Layout]>0,
a=dst[[1+p1Layout[[1,1]];;16+p1Layout[[1,1]],2;;25]];
conditionlabels=MapThread[StringJoin[#1,If[#1!=""&&#2!="",";",""],StringReplace[#2,"\n"->"+"]]&,{conditionlabels,a},2];
];

allplatesconditionlabels=<|#->conditionlabels&/@allBbarcodesDST|>;
allplatesconditionlabelsC=allplatesconditionlabels;
oneaddition[n_]:=Module[{},
Print["****** ",additions[[n]]<>".csv"," ******"];
wellvolume=wellvolumes[[n]];
Print["Well volume = ",wellvolume," ul."];
a=dst[[1+tables[[n]];;16+tables[[n]],2;;25]];
If[Count[a,s_String/;s!="",2]==0, (*4.12 after python*)
Print["Skipping empty Plate Map..."];
Print["Saving: ",Export[additions[[n]]<>".csv",StringRiffle[csv0,"\n",","],"Text"]]; 
Return[];
];


b={towellname[##],a[[##]]}&@@@Position[a,x_String/;x!=""];
c=Flatten[b/.{well_String,treat_String}:>({well,#}&/@StringSplit[treat,{"+",",",";","\n"}]),1];
If[(comp=Complement[Union[c[[All,2]]],labellists[[n,All,1]]])!={},Throw["Error: Missing label definition for: "<>comp]];
dosing=!MissingQ[ctables[tables[[n]]]];
If[dosing,
ac=dst[[1+ctables[tables[[n]]];;16+ctables[tables[[n]]],2;;25]];
cc=<|{towellname[##]->ac[[##]]}&@@@Position[ac,x_/;NumberQ[x]&&x>0] //Flatten|>;(*no splitting here*)
];


Print["Dispensing from: "<>barcodesSRC[[n]]];
activesrc=barcodesSRC[[n]];
With[{dst=#},
Print["Dispensing into: "<>dst];

platemapwithcompounds=a/.x_String:>StringRiffle[StringSplit[x,{"+",",",";","\n"}]/.compoundlabels[[n]],"+"];
allplatesconditionlabels[dst]=MapThread[StringJoin[#1,If[#1!=""&&#2!="",";",""],StringReplace[#2,"\n"->"+"]]&,{allplatesconditionlabels[dst],a},2];
allplatesconditionlabelsC[dst]=MapThread[StringJoin[#1,If[#1!=""&&#2!="",";",""],StringReplace[#2,"\n"->"+"]]&,{allplatesconditionlabelsC[dst],platemapwithcompounds},2];


Off[StringMatchQ::nonopt];
If[dosing,
l=#[[1]]->#[[2]]&/@labellists[[n]];
cfilt=Cases[c,_?(!MissingQ[cc[#[[1]]]]&)];
labelsdecoded = cfilt/.{x_String,y_}:>{x,y/.l}; (*250402 fix for using well labels as labels*)
diluteError[{#[[2]],#[[1]]/.cc}]&/@labelsdecoded;
cl={#[[1]],dilute[{#[[2]],#[[1]]/.cc}]}&/@labelsdecoded;
,
diluteError[#[[{2,3}]]]&/@labellists[[n]];
l=#[[1]]->dilute[#[[{2,3}]]]&/@labellists[[n]];
l=(#[[1,1]]->Flatten[#[[All,2]],1])&/@GatherBy[l,#[[1]]&];(*mod for multi compound per label*)
labelsdecoded = c/.{x_String,y_}:>{x,y/.l}; (*250402 fix for using well labels as labels*)
cl=Flatten[labelsdecoded/.{x_,y_}:>({x,#}&/@Partition[y,2]),1];(*mod for multi compound per label*)
];

d=Gather[Flatten[#,1]&/@cl,#1[[2]]==#2[[2]]&];
lastpos={1,1};
optim=reorderdst/@reordersrc[d];
optim=chooseSrc/@Flatten[optim,1];
optim=GatherBy[optim,#[[2]]&];

v={#[[1,2]]/.srcName[barcodesSRC[[n]]],barcodesSRC[[n]]<>"_"<>#[[1,2]],Total@#[[All,3]]/1000//N}&/@optim;
Print["Volumes used:"];
Print[StringRiffle[#,"   "]]&/@v;
totaluse=Append[totaluse,v];

(*
Print[ListLinePlot[Reverse[torowcol[#]]&/@#[[All,1]],PlotMarkers->{Automatic, Small},PlotRange->{{1,24},{1,16}}]&/@optim];
*)

csv=Flatten[{barcodesSRC[[n]],#[[1,2]],dst,#[[All,{3,1}]]}]&/@optim;


(*Print[csv];*)
If[additions0==additions[[n]],csv=Join[csv0,csv];Print["Merging picklist with previous."]];
additions0=additions[[n]];
csv0=csv;

]&/@If[barcodesDST[[n]]=={"*"},allBbarcodesDST,barcodesDST[[n]]];
Print["Saving: ",Export[additions[[n]]<>".csv",StringRiffle[csv,"\n",","],"Text"]];
]
csv0={};
additions0="";
totaluse={};
warnings=False;
Print[
Catch[
oneaddition/@Range[Length[additions]];
Print["****** Summary: Total Usage ******"];
Print[#[[1,1]],"\t\t",#[[1,2]],"\t",Total[#[[All,3]]],"\tremains (ul): ",volumes[#[[1,2]]]]&/@GatherBy[Flatten[totaluse,1],#[[2]]&];
src[[2;;,5]]=Values@volumes;
Export["picklist_after_dispense.xlsx","Sheets"->{worksheetnames[[1]]->src,worksheetnames[[2]]-> dst},"Rules"];

Print["***** Plateworks Tables ******"];

inventory={
MapIndexed[makeInventory[#1,SRCRack,#2[[1]]]&,Union[barcodesSRC]],
MapIndexed[makeInventory[#1,DSTRack,#2[[1]]]&,allBbarcodesDST]};
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
processtable=StringRiffle[Union[barcodesSRC],"\n"];
Print[processtable];
Export[ProcessSRC<>".txt",processtable,"Text"];
Print["***** Process_DST = "<>ProcessDST<>".txt"];
processtable=StringRiffle[allBbarcodesDST,"\n"];
Print[processtable];
Export[ProcessDST<>".txt",processtable,"Text"];
Export[f<>"_"<>#<>"_conditionslist labels.txt",StringRiffle[DeleteCases[Flatten[allplatesconditionlabels[#]],""],"\n"],"Text"]&/@allBbarcodesDST;
Export[f<>"_"<>#<>"_conditionslist compounds.txt",StringRiffle[DeleteCases[Flatten[allplatesconditionlabelsC[#]],""],"\n"],"Text"]&/@allBbarcodesDST;

mergtableperplate[plate_]:=Cases[Flatten[MapIndexed[{towellname@@#2,plate,#1}&, allplatesconditionlabels[plate],{2}],1],_?(#[[3]]!=""&)];
mergtableperplateC[plate_]:=Cases[Flatten[MapIndexed[{towellname@@#2,plate,#1}&, allplatesconditionlabelsC[plate],{2}],1],_?(#[[3]]!=""&)];
Export[f<>"_table merge labels.txt",
StringRiffle[(Flatten[mergtableperplate/@allBbarcodesDST,1])/.x_String/;StringMatchQ[x,___~~","|"\""~~___]:>"\""<>StringReplace[x,"\""->"\"\""]<>"\"","\n",","],"Text"];
Export[f<>"_table merge compounds.txt",
StringRiffle[(Flatten[mergtableperplateC/@allBbarcodesDST,1])/.x_String/;StringMatchQ[x,___~~","|"\""~~___]:>"\""<>StringReplace[x,"\""->"\"\""]<>"\"","\n",","],"Text"];

If[warnings,"Finished with warnings. Revise source file if needed!","Picklists passed all checks. Ready."]
]]

Print["Press Enter to exit..."];InputString[];




(* ::Input:: *)
(**)
(**)


(* ::Input:: *)
(**)
(**)
