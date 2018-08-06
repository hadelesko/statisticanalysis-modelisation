%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% ZIELE DIESER FUNKTION %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%% Anwendung von zwei Methoden für die Varianzanalyse der gesamten Daten jeder Aktivität und nach jeden
%%% Variablen: Ncht parametrische Test(KRUSKALWALLIS TEST ANOVA ONE WAY),
%%% weil alle Daten sind nicht normalverteilt UND DANN mit standard
%%% ANOVA(ANOVA mit p_wert von FISHER) basierend auf die Approximation der Daten zur
%%% Normalverteilung mit dem zentralen Grenzwertsatz                                                 
%%% Ergebnisse Zusammenfassen und Vergleich von der Ergebnisse der 2
%%% Methoden%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function[nva,table, p stats ]= fanotest3()
clc; clear; close all; 
nva=[{'AccX[m/s^2]'} {'AccY[m/s^2]'} {'AccZ[m/s^2]'} {'GyrX[degree/sec]'} {'GyrY[degree/sec]'} {'GyrZ[degree/sec]'}... 
     {'MagX[uT]'} {'MagY[uT]'} {'MagZ[uT]'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
 
nvaz=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
%nvazm=['AccX' 'AccY' 'AccZ' 'GyrX' 'GyrY' 'GyrZ' 'MagX' 'MagY' 'MagZ' 'Q0' 'Q1' 'Q2' 'Q3' 'Vbat'];

refz=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
s_zaehler = 1;



%tfile= strcat('KR',act,'.mat');
%act= 'WALK';   group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'	'V111'	'V114'	'V137' 'V141'}; % for WALK
%act= 'PICK';   group=	{'V66'   'V100'	'V111'	'V114'	'V137'	'V141'};      %for PICK	
%act= 'GRAP';   group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V111'	'V114' 'V64'};  %for GRAP
%act= 'PLACE';  group=	{'V30'	'V51'	'V56'	'V60'	'V66'	'V111'	'V114' 'V64' 'V61' 'V141'};  %for PLACE
%act= 'PUSH';   group=	{'V51'	'V56'	'V60'	'V61'	'V64'	'V66'};  %for PUSH
act= 'MT';     group=    {'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'};  %for MT

%t=uigetfile;
t= strcat('KD',act,'.xls');
dimension=size(xlsread(t,1));%Dateidimension inklusive Beschriftung(Probanden oder Group in ersten Zeile)
m=dimension(1);   % minus 1 wenn die gruppe Zahlen sind weil die Zeile der Gruppe in Datei soll nicht in berechnung angewendet werden 
n=dimension(2);   % oder n=size(group); % Anzahl der Probanden. 

%m= input('Geben Sie die Anzahl der Datenzeile ein!  ');
%n= input('Geben Sie die Anzahl der Probanden ein!    ');

% prompt =  {'Geben Sie die untersuchte Aktivität ein!'};
% dlg_title='The input Variable';
% inputd   =inputdlg(prompt,dlg_title);
% activity =upper(inputd(1));
desfile=strcat('KR',act,'.xls');
% Der Name der gewählten Datei oben
folder_a=strcat('KR',act);
%folder_b=strcat('SK',act);
mkdir(folder_a);
for i=s_zaehler:14      %%% Da 14 Variable und jede Variable eine Seite
   d = xlsread(t,s_zaehler);%% Offnung der n_te Seite der Datei 
   datatest = d(1:m,1:n);
   
   [p,table,stats] = kruskalwallis(datatest,group,'on');%%% Nicht-parametrische Test  Anova
   
   b1=strcat('Varianzanalyse/',act,'/', nvaz(s_zaehler),'/Kruskal-Wallis One-Way ANOVA/', 'PWert=', num2str(p));
   b2=strcat('Untersuchung der Aktivität _', act, '_ bei den Probanden an der Variablen: ',nva(s_zaehler));
   b3={'Varianzanalyse:Kruskal-Wallis One-Way ANOVA TABLE/Variable'};
   %k1={b2};
   kb1=cellstr(b1);
   kb2=cellstr(b2);
   title(b1);% Beschriftung der ANOVA Graphik
   ylabel(strcat('Ranks of mean of[',nvaz(s_zaehler),']'));
   
   [fp,ftable] = anova1(datatest,group,'on');%%% Anova Fisher
   disp(b2); disp(b3); disp(table);
   xlswrite(desfile,kb1, s_zaehler,'b1');
   xlswrite(desfile,kb2, s_zaehler,'b2');
   xlswrite(desfile,b3, s_zaehler,'b3'); 
   xlswrite(desfile,nva(s_zaehler), s_zaehler,'G3'); % Untersuchte Variable
   
   %title(strcat('Varianzanalyse für**',act,': Kruskal-Wallis One-Way ANOVA :  ', 'p-Wert=', num2str(p)))
   title(strcat(act,'/ANOVA ONE WAY/Variable', nva(s_zaehler)));
   xlabel('Kommissionierer');
   ylabel(nva(s_zaehler));
   disp(strcat('Untersuchung der**', act, '** bei den Kommissionierer:',group, 'an der Variablen: ',nva(s_zaehler)));
   disp('Varianzanalyse: Kruskal-Wallis  One-Way ANOVA');

   if p>0.05
       decision=strcat('Da PWert ',' = ',num2str(p),'   >   0.05;_');
       b4=strcat('** ENTSCHEIDUNG für** ', nva(s_zaehler));
       b5= ' Hypothese H1 verwerfen! Differenz zwischen Probanden NICHT SIGNIFIKANT';
       disp(b4); disp(b5);
       b6= strcat(decision,'  ',b4);
       b5= {b5};
       b6= cell(b6); 
       
       xlswrite(desfile,b6 ,s_zaehler,'b9');
       xlswrite(desfile,b5 ,s_zaehler,'b10');
       
       xlabel('Kommissionierer! Differenz nicht signifikant');
   else
      decision=strcat('Da PWert ',' = ',num2str(p),'   <  0.05 ;');
      b4=strcat('** ENTSCHEIDUNG für** ', nva(s_zaehler));
      b5='Hypothese H0 verwerfen! Differenz zwischen Probanden SIGNIFIKANT. ';
      disp(b4); disp(b5);
      b6= strcat(decision,'  ',b4);
      b5= {b5};
      b6= cell(b6);
      xlabel('Probanden/Kommissionier! Differenz zwischen Probanden ist signifikant.');
      
      xlswrite(desfile,b6,s_zaehler ,'b9');
      xlswrite(desfile,b5 ,s_zaehler,'b10');
   end
   
   %%%%% Multiple Comparison of mean of Variable
   
   multcompare(stats,'display','on');%%% Multiple Comparison of statistics aus Kruskalwallis ANOVA
   ylabel('Kommissionierer');
   title(strcat('Multiple Comparison of mean ranks of Variable[',nvaz(s_zaehler),']'));
   
   %%%% Test Statistik Kruskalwallis in Excelblatt schreiben 
   
   xlswrite(desfile,table, s_zaehler,'b4');%Kopie von Anova Table
   xlswrite(desfile,{'Untersuchte Videos'},s_zaehler,'d21');
   xlswrite(desfile, group,s_zaehler,'b22');% Untersuchte Gruppe

   %%%% Test Statistik Fisher Anova in Excelblatt schreiben
   xlswrite(desfile, {'ANOVA TABLE FISHER'},s_zaehler, 'd12');
   xlswrite(desfile,ftable, s_zaehler,'b13');%Kopie von Anova Table
   %xlswrite(desfile,fstatistics, s_zaehler,'b28');%%% Kopie von Statististics von ANOVA 
   %save(tfile, nva(s_zaehler));
   %save(tfile, fstats);
   if p>0.05
       df=strcat('PWert ',' =',num2str(fp), '>0.05', '_Hypothese H1 verwerfen!');
       decision={df};
       xlswrite(desfile,decision,s_zaehler,'B18');
       xlswrite(desfile,{'Differenz zwischen Kommissionierern ist NICHT SIGNIFIKANT'},s_zaehler,'B19');
   else
       df=strcat('PWert ',' =',num2str(fp), '<0.05', '_Hypothese H0 verwerfen!');
       decision={df};
       xlswrite(desfile,decision,s_zaehler,'B18')
       xlswrite(desfile,{'Differenz zwischen Kommissionierern ist SIGNIFIKANT'},s_zaehler,'B19'); 
   end

   
   %%%%%%% GRAPHIK SPEICHERN
   figurelist=findobj('type','figure');
   for j=1:numel(figurelist)
        saveas(figurelist(j),fullfile(folder_a,['figure' num2str(figurelist(j)) '.fig']));% speichert in Format.fig 
        saveas(figurelist(j),fullfile(folder_a,['figure' num2str(figurelist(j)) '.bmp']));  %speichert in Format .bmp 
   end
   
   s_zaehler=s_zaehler+1;
end
close(findobj('type','figure'));
   
%%%% ZUSAMMENFASSUNG DER VARIANZANALYSE FÜR DIE AKTIVITÄT 

zusf=strcat('AKTIVITÄT :', act,':ZUSAMMENFASSUNG DER VARIANZANALYSE');
zusammenfassung={zusf};
tsheet='zusammenfassung';
xlswrite(desfile,zusammenfassung,tsheet, 'F2');%%% Überschrift der Seite
xlswrite(desfile,nvaz,tsheet, 'B3');%%% Die Variablen
xlswrite(desfile,{'VARIABLE'},tsheet,'A3');
xlswrite(desfile,{'p_Wert   - aus KRUSKALWALLIS ANOVA'},tsheet,'A4');
xlswrite(desfile,{'p_Wert   - aus FISHER ANOVA'},tsheet,'A6');
xlswrite(desfile,{'HYPOTHESE  *** Kruskalwallis'},tsheet,'A5');
xlswrite(desfile,{'HYPOTHESE - ANOVA- Fisher '},tsheet,'A7');
xlswrite(desfile,{'HYPOTHESE - ANOVA- Fisher '},tsheet,'A7');
xlswrite(desfile,{'UNTERSUCHTE VIDEOS'},tsheet,'E12');
xlswrite(desfile,group,tsheet,'B13');
xlswrite(desfile,{'HYPOTHESE= 1; bedeutet H0 verwerfen! dh.**Differenz zwischen Kommissionierern ist  SIGNIFIKANT  '},tsheet,'B9');
xlswrite(desfile,{'HYPOTHESE= 0; bedeutet H1 verwerfen! dh.**Differenz zwischen Kommissionierern ist  NICHT SIGNIFIKANT'},tsheet,'B10');
w=1;
for x=w:14
    refzk='G5';
    refzf='G14';
    %%% Bildung der Referenze  oder Zellenposition 
    drefzk= strcat(refz(w),num2str(4));
    drefzf= strcat(refz(w),num2str(6));
    hrefzk= strcat(refz(w),num2str(5));
    hrefzf= strcat(refz(w),num2str(7));
    %%% P_Wert der Varianzanalyse lesen und kopieren in der Seite der
    %%% Zusamenfassung
    p_wk=xlsread(desfile,w,refzk);
    p_wf=xlsread(desfile,w,refzf); 
    xlswrite(desfile,p_wk,tsheet,drefzk);
    xlswrite(desfile,p_wf,tsheet,drefzf);
    
    %%% Welche Hypothese besagt den  p_Wert und ihre Bedeutung ? 
    if p_wk > 0.05
       xlswrite(desfile,'0',tsheet,hrefzk); %%% H0 annehmen
    else
       xlswrite(desfile,'1',tsheet,hrefzk); %%% H1 annehmen
    end
    if p_wf > 0.05
       xlswrite(desfile,'0',tsheet,hrefzf); %%% H0 annehmen: Differenz nicht Signifikant
    else
       xlswrite(desfile,'1',tsheet,hrefzf); %%% H1 annehmen: Differenz Signifikant
    end
    w=w+1; 
    
end
copyfile(desfile,folder_a);
end
%%
%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%This Function ist to calculate the Anova Fisher for the 14 Variable and
%%%to make the corresponding graphics
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%              function  sano.m                   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function[nvaz act fp ftable]= sanao()
clc; clear; close all; 
nvaz=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
%nvaz=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
s_zaehler=1;
%act= 'WALK';   group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'	'V111'	'V114'	'V137' 'V141'}; % for WALK
act= 'PICK';   group=	{'V66'  'V100'	'V111'	'V114'	'V137'	'V141'};      %for PICK	
%act= 'GRAP';   group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V111'	'V114' 'V64'};  %for GRAP
%act= 'PLACE';  group=	{'V30'	'V51'	'V56'	'V60'	'V66'	'V111'	'V114' 'V64' 'V61' 'V141'};  %for PLACE
%act= 'PUSH';   group=	{'V51'	'V56'	'V60'	'V61'	'V64'	'V66'};  %for PUSH
%act= 'MT';     group=    {'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'};  %for MT
t= strcat('KD',act,'.xls');
folder_b= strcat('AF',act);
mkdir(folder_b);
for ij=s_zaehler:14      %%% Da 14 Variable und jede Variable eine Seite
   d = xlsread(t,s_zaehler);%% Offnung der n_te Seite der Datei 
   %dataa = d(:,:);
   [fp,ftable,fstats] = anova1(d,group,'on');%%% Anova Fisher
   title(strcat('Multiple Comparison of mean of Rank of  Variable[',nvaz(s_zaehler),']')); 
   multcompare(fstats,'display','on');%%% Multiple Comparison of statistics aus Fisher ANOVA
   ylabel('Kommissionierer'); title(strcat('Multiple Comparison of mean of Variable[',nvaz(s_zaehler),']')); 
   
   figurelist=findobj('type','figure');
   for j=1:numel(figurelist)
        saveas(figurelist(j),fullfile(folder_b,['figure' num2str(figurelist(j)) '.fig']));% speichert in Format.fig 
        saveas(figurelist(j),fullfile(folder_b,['figure' num2str(figurelist(j)) '.bmp']));  %speichert in Format .bmp 
   end
   s_zaehler=s_zaehler+1;
   
%%%%%% 
end
close(findobj('type','figure'));
end

%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%       PRUEFFUNGQPLOT        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% FONCTION NORMALITÄT PRÜFUNG %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%% Prüfung der Normalverteilung anhand qqplot
%%% Datenfile sind : KM+activity + .xlsx
%%% Seite der Daten: 'activity'
function[activity sfile ]=PRUEFFUNGQPLOT()
close all;
%sfile= uigetfile;
%sf=[{'KMWALK.xlsx'} {'KMPICK.xlsx'} {'KMGRAP.xlsx'} {'KMPLACE.xlsx'} {'KMMT.xlsx'}  {'KMPUSH.xlsx'}];
%activity =[{'WALK'} {'PICK'} {'GRAP'} {'PLACE'} {'MT'} {'PUSH'}];

% nva=[{'AccX[m/s^2]'} {'AccY[m/s^2]'} {'AccZ[m/s^2]'} {'GyrX[degree/sec]'} {'GyrY[degree/sec]'} {'GyrZ[degree/sec]'}... 
     %{'MagX[uT]'} {'MagY[uT]'} {'MagZ[uT]'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
nva=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
d='Normalprüfung'; dkf='Dichtkurve';
mkdir(d);
mkdir(dkf);
z=1;zz=1;
v=1;
activity='WALK'; 
% activity='PICK';
% activity='GRAP';
% activity='PLACE';
% activity='PUSH';
% activity='MT';
sfile = strcat('KM', activity,'.xlsx');
s= xlsread(sfile,activity);
h=figure('Name',activity);

%Test auf die Normaltät: One-sample Kolmogorov-Smirnov test und mit qq-plot
for i=z:14
       H = kstest(s(:,z));  %%% Test auf die Normaltät: One-sample Kolmogorov-Smirnov test
       hypothese=strcat(nva(z),'[h= ',num2str(H),' ]');
       subplot(5,3,z);qqplot(s(:,z));legend(hypothese,'Location', 'SouthEast');
       
       z=z+1;
end
saveas(h,fullfile(d,[activity '.fig']));% speichert in Format.fig 
saveas(h,fullfile(d,[activity '.pdf']));
saveas(h,fullfile(d,[activity '.png']));
saveas(h,fullfile(d,[activity '.jpeg']));
saveas(h,fullfile(d,[activity '.bmp']));

%%% Normalitätprüfung mit Dichte Kurve
dk= figure('Name',strcat('Dichte Kurve_',activity));
for j=zz:14
 x=s(:,zz);   
 mu=mean(x);
 sig=std(x);
 pdfn=normpdf(x,mu,sig);
 subplot(5,3,zz); plot(x,pdfn);xlabel(nva(zz)); ylabel('Pr.density');
 zz=zz+1;
end
saveas(dk,fullfile(dkf,[activity '.fig']));% speichert in Format.fig 
saveas(dk,fullfile(dkf,[activity '.pdf']));
saveas(dk,fullfile(dkf,[activity '.png']));
saveas(dk,fullfile(dkf,[activity '.jpeg']));
saveas(dk,fullfile(dkf,[activity '.bmp']));
copyfile(dkf,d); %%% Copy the folder dkf in the folder d 
%%%%
end
%%
%%% Hier kombinieren wir die Probanden für bestimmte Aktivität nach
%%% bestimmte Variable. Die Quelledatei ist die Datei mit Präfix 'KM' und
%%% die Destination Datei ist die Datei mit Präfix 'KD'
%%% Da die Häufigkeit der Aktivität ist nicht alle Prozess nicht gleich
%%% ist,müssen wir zuerst die Maximale Häufigkeit der Aktivität aus
%%% betrachteten Probanden identifizieren( beim lesen SOURCE File der Akivität(Datei mit
%%% Präfix 'KM') [entspricht datenzeile]),
%%% dann alle die Dimension (in Zeile) auf jeder Seite der 'KM'-Datei
%%% uniformisieren. Die Uniformisierung ist besteht hier die leeren Zeile auf
%%% Probendenseite mit Null'0'auffüllen Seite, damit die Datenlänge(Zeile)
%%% auf jede Seite gleich wird.
%%% Mit der Funktion [vorhyooo3] ist der Skript [Meancombinefile] automatisch
%%% generated. With these skript wird einen neuen skript automatisch
%%% generated%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function[ansheet,nvideo]=vorhyooo3median( )
% refz={'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'};
% 'Geben Sie die untersuchte Aktivität  als konstante ein!'
dsr= {'desrange = ' 'B2'};
colref = {'refz=[' 'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o' '];'};
prompt =  {'Geben Sie die untersuchte Aktivität ein!'};
          
dlg_title='The input Variable';
inputd   =inputdlg(prompt,dlg_title);
activity =upper(inputd(1));

adfile = strcat('KMD', activity, '.xls');
sfile= strcat('MD',activity, '.xlsx'); % The  excel workbook with activity’s Mean by every Video
nvideo = input('Geben Sie bitte die Anzahl der untersuchten Videos ein:   ');
varz=1;
disp(strcat ('%%%  switch ',' _ ', activity));
disp('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%');
disp(strcat(upper('% Generated Command to copy the  _'), upper('._ Variable')));
% Destination sheet
disp(strcat(' % case ','  ',num2str(varz) ));
disp(upper('%%% Hinweise: Ab hier können sie den generated Code kopieren und löschen und füge die einzelzitatfürhung!!!'))
disp(colref);
disp(strcat('sfile = ',sfile, ';  %"', sfile,'" in EinzelZitatmarken einsetzen'));
disp(strcat('dfile = ',adfile,';  %"activity.xls" in Einzelzitatmarken ensezten'));
disp('lendata   = size(xlsread(sfile,1));');
disp('datenzeile = lendata(1);');
disp('refn   =  1 ;');
disp(dsr);
disp('for  i =refn:14')
disp('Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));');
disp('dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung');
disp(strcat('xlswrite(','dfile',',['));     % The beginning of the Synthax to write the  matrix of the combined sheets
ansheet=1;
    for i=ansheet:nvideo-1		 % Count the sheets in the original Worksbooks
        sheet=num2str(ansheet);
        %The beginning of the Syntax to write the sheets that will be combined. Here we add ',' after 'Range)' to write  the
        %data in  column and '...' to signify to  Matlab to stay on the same line although the command goes to the next line
        disp(strcat('xlsread(','sfile',',',sheet,',','Range',')',',','...'));
        ansheet=ansheet+1;
    end
    
    disp(strcat('xlsread(','sfile',',', num2str(nvideo),',Range',')], dsheet , desrange);'));
    disp('refn=refn+1;');
    disp('end');
    disp('%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%');
    disp('% Kopiere die oben angeszeigte Code von "dsheet=... bis dsheet,desrange)" und führen Sie das aus!!')
    disp(upper('% Vielen Dank für ihre Aufmerksamkeit.%'));
    %varz=varz+1;
end 
%end
%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%% Hier kombinieren wir die Probanden für bestimmte Aktivität nach
%%% bestimmte Variable. Die Quelledatei ist die Datei mit Präfix 'KM' und
%%% die Destination Datei ist die Datei mit Präfix 'KD'
%%% Da die Häufigkeit der Aktivität ist nicht alle Prozess nicht gleich
%%% ist,müssen wir zuerst die Maximale Häufigkeit der Aktivität aus
%%% betrachteten Probanden identifizieren( beim lesen SOURCE File der Akivität(Datei mit
%%% Präfix 'KM') [entspricht datenzeile]),
%%% dann alle die Dimension (in Zeile) auf jeder Seite der 'KM'-Datei
%%% uniformisieren. Die Uniformisierung ist besteht hier die leeren Zeile auf
%%% Probendenseite mit Null'0'auffüllen Seite, damit die Datenlänge(Zeile)
%%% auf jede Seite gleich wird.
%%% Mit der Funktion [vorhyooo3] ist der Skript [Meancombinefile] automatisch
%%% generated. With these skript wird einen neuen skript automatisch
%%% generated%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function[ansheet,nvideo]=vorhyooo3mean( )
% refz={'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'};
% 'Geben Sie die untersuchte Aktivität  als konstante ein!'
dsr= {'desrange = ' 'B2'};
colref = {'refz=[' 'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o' '];'};
prompt =  {'Geben Sie die untersuchte Aktivität ein!'};
          
dlg_title='The input Variable';
inputd   =inputdlg(prompt,dlg_title);
activity =upper(inputd(1));

adfile = strcat('KM', activity, '.xls');
sfile= strcat('KD',activity, '.xlsx'); % The  excel workbook with activity’s Mean by every Video
nvideo = input('Geben Sie bitte die Anzahl der untersuchten Videos ein:   ');
varz=1;
disp(strcat ('%%%  switch ',' _ ', activity));
disp('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%');
disp(strcat(upper('% Generated Command to copy the  _'), upper('._ Variable')));
% Destination sheet
disp(strcat(' % case ','  ',num2str(varz) ));
disp(upper('%%% Hinweise: Ab hier können sie den generated Code kopieren und löschen und füge die einzelzitatfürhung!!!'))
disp(colref);
disp(strcat('sfile = ',sfile, ';  %"', sfile,'" in EinzelZitatmarken einsetzen'));
disp(strcat('dfile = ',adfile,';  %"activity.xls" in Einzelzitatmarken ensezten'));
disp('lendata   = size(xlsread(sfile,1));');
disp('datenzeile = lendata(1);');
disp('refn   =  1 ;');
disp(dsr);
disp('for  i =refn:14')
disp('Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));');
disp('dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung');
disp(strcat('xlswrite(','dfile',',['));     % The beginning of the Synthax to write the  matrix of the combined sheets
ansheet=1;
    for i=ansheet:nvideo-1		 % Count the sheets in the original Worksbooks
        sheet=num2str(ansheet);
        %The beginning of the Syntax to write the sheets that will be combined. Here we add ',' after 'Range)' to write  the
        %data in  column and '...' to signify to  Matlab to stay on the same line although the command goes to the next line
        disp(strcat('xlsread(','sfile',',',sheet,',','Range',')',',','...'));
        ansheet=ansheet+1;
    end
    
    disp(strcat('xlsread(','sfile',',', num2str(nvideo),',Range',')], dsheet , desrange);'));
    disp('refn=refn+1;');
    disp('end');
    disp('%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%');
    disp('% Kopiere die oben angeszeigte Code von "dsheet=... bis dsheet,desrange)" und führen Sie das aus!!')
    disp(upper('% Vielen Dank für ihre Aufmerksamkeit.%'));
    %varz=varz+1;
end 
%%
%%% Hier kombinieren wir die Probanden für bestimmte Aktivität nach
%%% bestimmte Variable. Die Quelledatei ist die Datei mit Präfix 'KM' und
%%% die Destination Datei ist die Datei mit Präfix 'KD'
%%% Da die Häufigkeit der Aktivität ist nicht alle Prozess nicht gleich
%%% ist,müssen wir zuerst die Maximale Häufigkeit der Aktivität aus
%%% betrachteten Probanden identifizieren( beim lesen SOURCE File der Akivität(Datei mit
%%% Präfix 'KM') [entspricht datenzeile]),
%%% dann alle die Dimension (in Zeile) auf jeder Seite der 'KM'-Datei
%%% uniformisieren. Die Uniformisierung ist besteht hier die leeren Zeile auf
%%% Probendenseite mit Null'0'auffüllen Seite, damit die Datenlänge(Zeile)
%%% auf jede Seite gleich wird.
%%% Mit der Funktion [vorhyooo3] ist der Skript [Meancombinefile] automatisch
%%% generated. With these skript wird einen neuen skript automatisch
%%% generated%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function[ansheet,nvideo]=vorhyooo3median( )
% refz={'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'};
% 'Geben Sie die untersuchte Aktivität  als konstante ein!'
dsr= {'desrange = ' 'B2'};
colref = {'refz=[' 'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o' '];'};
prompt =  {'Geben Sie die untersuchte Aktivität ein!'};
          
dlg_title='The input Variable';
inputd   =inputdlg(prompt,dlg_title);
activity =upper(inputd(1));

adfile = strcat('KMD', activity, '.xls');
sfile= strcat('MD',activity, '.xlsx'); % The  excel workbook with activity’s Mean by every Video
nvideo = input('Geben Sie bitte die Anzahl der untersuchten Videos ein:   ');
varz=1;
disp(strcat ('%%%  switch ',' _ ', activity));
disp('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%');
disp(strcat(upper('% Generated Command to copy the  _'), upper('._ Variable')));
% Destination sheet
disp(strcat(' % case ','  ',num2str(varz) ));
disp(upper('%%% Hinweise: Ab hier können sie den generated Code kopieren und löschen und füge die einzelzitatfürhung!!!'))
disp(colref);
disp(strcat('sfile = ',sfile, ';  %"', sfile,'" in EinzelZitatmarken einsetzen'));
disp(strcat('dfile = ',adfile,';  %"activity.xls" in Einzelzitatmarken ensezten'));
disp('lendata   = size(xlsread(sfile,1));');
disp('datenzeile = lendata(1);');
disp('refn   =  1 ;');
disp(dsr);
disp('for  i =refn:14')
disp('Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));');
disp('dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung');
disp(strcat('xlswrite(','dfile',',['));     % The beginning of the Synthax to write the  matrix of the combined sheets
ansheet=1;
    for i=ansheet:nvideo-1		 % Count the sheets in the original Worksbooks
        sheet=num2str(ansheet);
        %The beginning of the Syntax to write the sheets that will be combined. Here we add ',' after 'Range)' to write  the
        %data in  column and '...' to signify to  Matlab to stay on the same line although the command goes to the next line
        disp(strcat('xlsread(','sfile',',',sheet,',','Range',')',',','...'));
        ansheet=ansheet+1;
    end
    
    disp(strcat('xlsread(','sfile',',', num2str(nvideo),',Range',')], dsheet , desrange);'));
    disp('refn=refn+1;');
    disp('end');
    disp('%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%');
    disp('% Kopiere die oben angeszeigte Code von "dsheet=... bis dsheet,desrange)" und führen Sie das aus!!')
    disp(upper('% Vielen Dank für ihre Aufmerksamkeit.%'));
    %varz=varz+1;
end
%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% ZIELE DIESER FUNKTION %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%% Anwendung von zwei Methoden für die Varianzanalyse der gesamten Daten jeder Aktivität und nach jeden
%%% Variablen: Ncht parametrische Test(KRUSKALWALLIS TEST ANOVA ONE WAY),
%%% weil alle Daten sind nicht normalverteilt UND DANN mit standard
%%% ANOVA(ANOVA mit p_wert von FISHER) basierend auf die Approximation der Daten zur
%%% Normalverteilung mit dem zentralen Grenzwertsatz                                                 
%%% Ergebnisse Zusammenfassen und Vergleich von der Ergebnisse der 2
%%% Methoden%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function[nva,table, p stats ]= fanotest3med()
clc;clear;close all; 
nva=[{'AccX[m/s^2]'} {'AccY[m/s^2]'} {'AccZ[m/s^2]'} {'GyrX[degree/sec]'} {'GyrY[degree/sec]'} {'GyrZ[degree/sec]'}... 
     {'MagX[uT]'} {'MagY[uT]'} {'MagZ[uT]'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
 
nvaz=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
%nvazm=['AccX' 'AccY' 'AccZ' 'GyrX' 'GyrY' 'GyrZ' 'MagX' 'MagY' 'MagZ' 'Q0' 'Q1' 'Q2' 'Q3' 'Vbat'];

refz=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
s_zaehler = 1;

%act= 'WALK';  group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'	'V111'	'V114'	'V137' 'V141'}; % for WALK
%act= 'PICK';  group=	{'V66'  'V100'	'V111'	'V114'	'V137'	'V141'};      %for PICK	
%act= 'GRAP';  group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'  'V111'};  %for GRAP
act= 'PLACE'; group=	{'V30'	'V51'	'V56'	'V60'   'V61'  'V64'   'V66'    'V111'	'V114' 'V141'};  %for PLACE
%act= 'PUSH';  group=	{'V51'	'V56'	'V60'	'V61'	'V64'	'V66'};  %for PUSH
%act= 'MT';   group=   {'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'};  %for MT

t=strcat('KMD',upper(act),'.xls'); %t=uigetfile;
dimension=size(xlsread(t,1));%Dateidimension inklusive Beschriftung(Probanden oder Group in ersten Zeile)
m=dimension(1);   % minus 1 wenn die gruppe Zahlen sind weil die Zeile der Gruppe in Datei soll nicht in berechnung angewendet werden 
n=dimension(2);   % oder n=size(group); % Anzahl der Probanden. 

%m= input('Geben Sie die Anzahl der Datenzeile ein!  ');
%n= input('Geben Sie die Anzahl der Probanden ein!    ');


desfile=strcat('KRMD',act,'.xls');
% Der Name der gewählten Datei oben
folder_a=strcat('KRMD',act);
mkdir(folder_a);

for i=s_zaehler:14      %%% Da 14 Variable und jede Variable eine Seite
   d = xlsread(t,s_zaehler);%% Offnung der n_te Seite der Datei 
   datatest = d(1:m,1:n);
   
   [p,table,stats] = kruskalwallis(datatest,group,'on');%%% Nicht-parametrische Test  Anova
   
   b1=strcat('Varianzanalyse/',act,'/', nvaz(s_zaehler),'/Kruskal-Wallis One-Way ANOVA/', 'PWert=', num2str(p));
   b2=strcat('Untersuchung der Aktivität _', act, '_ bei den Probanden an der Variablen: ',nva(s_zaehler));
   b3={'Varianzanalyse:Kruskal-Wallis One-Way ANOVA TABLE/Variable'};
   k1={b2};
   kb1=cellstr(b1);
   kb2=cellstr(b2); 
   title(b1);% Beschriftung der ANOVA Graphik
   ylabel(strcat('Ranks of mean of[',nvaz(s_zaehler),']'));
   
   [fp,ftable] = anova1(datatest,group,'on');%%% Anova Fisher
   disp(b2); disp(b3); disp(table);
   xlswrite(desfile,b1, s_zaehler,'b1');
   xlswrite(desfile,kb2, s_zaehler,'b2');
   xlswrite(desfile,b3, s_zaehler,'b3'); 
   xlswrite(desfile,nva(s_zaehler), s_zaehler,'G3'); % Untersuchte Variable
   
   %title(strcat('Varianzanalyse für**',act,': Kruskal-Wallis One-Way ANOVA :  ', 'p-Wert=', num2str(p)))
  
   title(strcat(act,'/ANOVA ONE WAY/Variable', nva(s_zaehler)));
   xlabel('Kommissionierer');
   ylabel(nva(s_zaehler));
   disp(strcat('Untersuchung der**', act, '** bei den Kommissionierer:',group, 'an der Variablen: ',nva(s_zaehler)));
   disp('Varianzanalyse: Kruskal-Wallis  One-Way ANOVA');

   if p>0.05
       decision=strcat('Da PWert ',' = ',num2str(p),'   >   0.05;_');
       b4=strcat('** ENTSCHEIDUNG für** ', nva(s_zaehler));
       b5= ' Hypothese H1 verwerfen! Differenz zwischen Probanden NICHT SIGNIFIKANT';
       disp(b4); disp(b5);
       b6= strcat(decision,'  ',b4);
       b5= {b5};
       b6= cell(b6); 
       
       xlswrite(desfile,b6 ,s_zaehler,'b9');
       xlswrite(desfile,b5 ,s_zaehler,'b10');
       
       xlabel('Kommissionierer! Differenz nicht signifikant');
   else
      decision=strcat('Da PWert ',' = ',num2str(p),'   <  0.05 ;');
      b4=strcat('** ENTSCHEIDUNG für** ', nva(s_zaehler));
      b5='Hypothese H0 verwerfen! Differenz zwischen Probanden SIGNIFIKANT. ';
      disp(b4); disp(b5);
      b6= strcat(decision,'  ',b4);
      b5= {b5};
      b6= cell(b6);
      xlabel('Probanden/Kommissionier! Differenz zwischen Probanden ist signifikant.');
      
      xlswrite(desfile,b6,s_zaehler ,'b9');
      xlswrite(desfile,b5 ,s_zaehler,'b10');
   end
   
   %%%%% Multiple Comparison of mean of Variable
   
   multcompare(stats,'display','on');%%% Multiple Comparison of statistics aus Kruskalwallis ANOVA
   ylabel('Kommissionierer');
   title(strcat('Multiple Comparison of mean ranks of Variable[',nvaz(s_zaehler),']'));
   
   %%%% Test Statistik Kruskalwallis in Excelblatt schreiben 
   
   xlswrite(desfile,table, s_zaehler,'b4');%Kopie von Anova Table
   xlswrite(desfile,{'Untersuchte Videos'},s_zaehler,'d21');
   xlswrite(desfile, group,s_zaehler,'b22');% Untersuchte Gruppe

   %%%% Test Statistik Fisher Anova in Excelblatt schreiben
   xlswrite(desfile, {'ANOVA TABLE FISHER'},s_zaehler, 'd12');
   xlswrite(desfile,ftable, s_zaehler,'b13');%Kopie von Anova Table
   %xlswrite(desfile,fstatistics, s_zaehler,'b28');%%% Kopie von Statististics von ANOVA 
   %save(tfile, nva(s_zaehler));
   %save(tfile, fstats);
   if p>0.05
       df=strcat('PWert ',' =',num2str(fp), '>0.05', '_Hypothese H1 verwerfen!');
       decision={df};
       xlswrite(desfile,decision,s_zaehler,'B18');
       xlswrite(desfile,{'Differenz zwischen Kommissionierern ist NICHT SIGNIFIKANT'},s_zaehler,'B19');
   else
       df=strcat('PWert ',' =',num2str(fp), '<0.05', '_Hypothese H0 verwerfen!');
       decision={df};
       xlswrite(desfile,decision,s_zaehler,'B18')
       xlswrite(desfile,{'Differenz zwischen Kommissionierern ist SIGNIFIKANT'},s_zaehler,'B19'); 
   end

   
   %%%%%%% GRAPHIK SPEICHERN
   figurelist=findobj('type','figure');
   for j=1:numel(figurelist)
        saveas(figurelist(j),fullfile(folder_a,['figure' num2str(figurelist(j)) '.fig']));% speichert in Format.fig 
        saveas(figurelist(j),fullfile(folder_a,['figure' num2str(figurelist(j)) '.bmp']));  %speichert in Format .bmp 
   end
   
   s_zaehler=s_zaehler+1;
end
close(findobj('type','figure'));
   
%%%% ZUSAMMENFASSUNG DER VARIANZANALYSE FÜR DIE AKTIVITÄT 

zusf=strcat('AKTIVITÄT :', act,':ZUSAMMENFASSUNG DER VARIANZANALYSE');
zusammenfassung={zusf};
tsheet='zusammenfassung';
xlswrite(desfile,zusammenfassung,tsheet, 'F2');%%% Überschrift der Seite
xlswrite(desfile,nvaz,tsheet, 'B3');%%% Die Variablen
xlswrite(desfile,{'VARIABLE'},tsheet,'A3');
xlswrite(desfile,{'p_Wert   - aus KRUSKALWALLIS ANOVA'},tsheet,'A4');
xlswrite(desfile,{'p_Wert   - aus FISHER ANOVA'},tsheet,'A6');
xlswrite(desfile,{'HYPOTHESE  *** Kruskalwallis'},tsheet,'A5');
xlswrite(desfile,{'HYPOTHESE - ANOVA- Fisher '},tsheet,'A7');
xlswrite(desfile,{'HYPOTHESE - ANOVA- Fisher '},tsheet,'A7');
xlswrite(desfile,{'UNTERSUCHTE VIDEOS'},tsheet,'E12');
xlswrite(desfile,group,tsheet,'B13');
xlswrite(desfile,{'HYPOTHESE= 1; bedeutet H0 verwerfen! dh.**Differenz zwischen Kommissionierern ist  SIGNIFIKANT  '},tsheet,'B9');
xlswrite(desfile,{'HYPOTHESE= 0; bedeutet H1 verwerfen! dh.**Differenz zwischen Kommissionierern ist  NICHT SIGNIFIKANT'},tsheet,'B10');
w=1;
for x=w:14
    refzk='G5';
    refzf='G14';
    %%% Bildung der Referenze  oder Zellenposition 
    drefzk= strcat(refz(w),num2str(4));
    drefzf= strcat(refz(w),num2str(6));
    hrefzk= strcat(refz(w),num2str(5));
    hrefzf= strcat(refz(w),num2str(7));
    %%% P_Wert der Varianzanalyse lesen und kopieren in der Seite der
    %%% Zusamenfassung
    p_wk=xlsread(desfile,w,refzk);
    p_wf=xlsread(desfile,w,refzf); 
    xlswrite(desfile,p_wk,tsheet,drefzk);
    xlswrite(desfile,p_wf,tsheet,drefzf);
    
    %%% Welche Hypothese besagt den  p_Wert und ihre Bedeutung ? 
    if p_wk > 0.05
       xlswrite(desfile,'0',tsheet,hrefzk); %%% H0 annehmen
    else
       xlswrite(desfile,'1',tsheet,hrefzk); %%% H1 annehmen
    end
    if p_wf > 0.05
       xlswrite(desfile,'0',tsheet,hrefzf); %%% H0 annehmen: Differenz nicht Signifikant
    else
       xlswrite(desfile,'1',tsheet,hrefzf); %%% H1 annehmen: Differenz Signifikant
    end
    w=w+1; 
    
end
copyfile(desfile,folder_a);
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%
%%%% jOIN THE VARIANCE ANALYSE Result by mean
%%%% The file name change by the mediandata analyse
%%%%  jOIN THE T-TEST Result
function[varsheet nvaz]=joinalltvarresult()
varsheet='zusammenfassung';
nvaz=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
xlswrite('allvarianzttest.xls',...
    [xlsread('KRMDWALK.xls','NTTESTWALK');...
    xlsread('KRMDPICK.xls','NTTESTPICK');...
    xlsread('KRMDGRAP.xls','NTTESTGRAP');...
    xlsread('KRMDPLACE.xls','NTTESTPLACE');...
    xlsread('KRMDMT.xls','NTTESTMT');...
    xlsread('KRMDPUSH.xls','NTTESTPUSH')], 'allactvar', 'C3');
xlswrite('allvarianzttest.xls' ,nvaz,'allactvar', 'C2');
xlswrite('allvarianzttest.xls',...
    [xlsread('KRMDWALK.xls',varsheet);...
    xlsread('KRMDPICK.xls',varsheet);...
    xlsread('KRMDGRAP.xls',varsheet);...
    xlsread('KRMDPLACE.xls',varsheet);...
    xlsread('KRMDMT.xls',varsheet);...
    xlsread('KRMDPUSH.xls',varsheet)], 'allactttest', 'C3');
xlswrite('allvarianzttest.xls' ,nvaz,'allactttest','C2');
end

function[sheet       az       ez        n_activity     dfile      sfile ]        =        Combdstat()
az= input('Geben Sie die Anfangzeile der Aktivität ein: '); %Aus Aktivitätsfile zu entnehmen 
ez= input('Geben Sie die Endezeile der Aktivität ein: '); %Aus Aktivitätsfile zu entnehmen
n_activity= ez-az+1; %Häufigkeit der Aktivität bei der Kommissionierung

disp('%% Wichtige Anweisungen: Folgen Sie den Anweisungen unten %%');
desfile= 'SD11.xls';
% Die Anweisung für die Ausführung von der erhaltene Befehle
disp(strcat(upper('% Zur Erstellung von Exceltabelle "'),desfile,upper(' "für die statistische Analyse: '), upper(' To Do:')));
disp('%1-) Kopieren Sie die oben angezeigte gesamte Kommandos von "xlswrite(...bis desheet,desrange)"'); 
disp('%2-) Fügen Sie den kopierten Befehl auf der Kommandozeile ein');
disp('%3-) Setzen Sie bitte {sstat.xls},{SD11.xls},{B7:O7},{Meanactivity},{B2} in Einzelzitatenmarken.');
disp('%4-) Führen Sie das direkt in Kommandos windows aus! ');
disp('% %');
disp(strcat('sfile= sstat.xls', ' ;'));
disp(strcat('dfile= SD11.xls', ' ;'));
disp(['Range ','= ', 'B7:O7 ;']); % The part or range of the the sheet where the desired statistics is 
disp(['dsheet ','=', ' Meanactivity ;']);%% Destination sheet
disp(['desrange',' = ','B2 ;']); %%% Destinaions Range in destination sheet
%Display the Part of the script that will be used to combine the data of the considered activity 
%for the global statistics of the activity
dfile={'SD11.xls'}; %%%%The destinations file that will contains the combined data of activity
disp(strcat('xlswrite(','dfile',',[')); % The beginning of the Synthax to write
% the matrix of the combined sheets disp(strcat('sfile = Day3.xls', ' ;'));%The source file that contains
% the combined data of activity of the AccY-axes

for j=az:ez-1
      sfile={'sstat.xls'};
      sheet=az; 
      % The beginning of the Synthax to write the sheets that will be combined y axes 
      disp(strcat('xlsread(','sfile',',', num2str(sheet),', Range',')',';')); 
      az=az+1; 
end
%This is the end of the matrix synthaxe.Die letzte Zeile der Befehl mit Destinationsblatt und Range in Excelblatt 

disp(strcat('xlsread(','sfile',',', num2str(ez),', Range',')], dsheet, desrange)')); % 

disp('% %');
disp(upper('% Vielen Dank für ihre Aufmerksamkeit.%')); 
%%% Den Kopf von SD11.xls vorbereiten
xlswrite('SD11.xls',[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'}...
{'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}],'Meanactivity','B1:O1'); 
end
%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%
%%
%%% Hier kombinieren wir die Probanden für bestimmte Aktivität nach
%%% bestimmte Variable. Die Quelledatei ist die Datei mit Präfix 'KM' und
%%% die Destination Datei ist die Datei mit Präfix 'KD'
%%% Da die Häufigkeit der Aktivität ist nicht alle Prozess nicht gleich
%%% ist,müssen wir zuerst die Maximale Häufigkeit der Aktivität aus
%%% betrachteten Probanden identifizieren( beim lesen SOURCE File der Akivität(Datei mit
%%% Präfix 'KM') [entspricht datenzeile]),
%%% dann alle die Dimension (in Zeile) auf jeder Seite der 'KM'-Datei
%%% uniformisieren. Die Uniformisierung ist besteht hier die leeren Zeile auf
%%% Probendenseite mit Null'0'auffüllen Seite, damit die Datenlänge(Zeile)
%%% auf jede Seite gleich wird.
%%% Mit der Funktion [vorhyooo3] ist der Skript [Meancombinefile] automatisch
%%% generated
%%%
function [sfile,dfile,lendata,datenzeile,desrange,Range]=Meancombinefile()

refz=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
activity='WALK'; 
% activity='PICK';
% activity='GRAP';
% activity='PLACE';
% activity='PUSH';
% activity='MT';
sfile =strcat('KM' , activity,'.xlsx');     
dfile =strcat('KD' , activity,'.xls');
%%% In diesem Teil ist für jede Aktivität konstant.Keine Änderung
lendata= size(xlsread(sfile,1));
datenzeile=lendata(1);%datenzeile=('Geben Sie die Anzahl der Zeile der Daten in Quelledatei:   ');
desrange = 'B2'; % 
refn=1;% Zähler der 14 Seiten zu schreiben

for i=refn:14
    Range=strcat(refz(refn),num2str(2),':',refz(refn),num2str(datenzeile));
    dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file          
    %%%% Bis hier keine Änderung
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%   Here we paste the outputskript of the functionvvorhyooo3.m 
    %%%  (only the part xlswrite(...[....],dsheet, desrange]
end
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    

%%%%%%% Exemple with WALK

refz=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
sfile ='KMWALK.xlsx';     
dfile ='KDWALK.xls'; 
%%% In diesem Teil ist für jede Aktivität konstant.Keine Änderung
lendata= size(xlsread(sfile,1));
datenzeile=lendata(1);%datenzeile=('Geben Sie die Anzahl der Zeile der Daten in Quelledatei:   ');
desrange = 'B2'; % 
refn=1;% Zähler der 14 Seiten zu schreiben

for i=refn:14
    Range=strcat(refz(refn),num2str(2),':',refz(refn),num2str(datenzeile));
    dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file          
    %%%% Bis hier keine Änderung
   
    xlswrite(dfile,[
    xlsread(sfile,1,Range),...
    xlsread(sfile,2,Range),...
    xlsread(sfile,3,Range),...
    xlsread(sfile,4,Range),...
    xlsread(sfile,5,Range),...
    xlsread(sfile,6,Range),...
    xlsread(sfile,7,Range),...
    xlsread(sfile,8,Range),...
    xlsread(sfile,9,Range),...
    xlsread(sfile,10,Range),...
    xlsread(sfile,11,Range),...
    xlsread(sfile,12,Range)],dsheet, desrange)


    refn=refn+1;
end
end
%%
 
%%%%%%% Exemple with PICK
 
 
refz = [    'b'    'c'    'd'    'e'    'f'    'g'    'h'    'i'    'j'...
    'k'    'l'    'm'    'n'    'o'    ];

    sfile ='KMPICK.xlsx';  %"KMPICK.xlsx" in EinzelZitatmarken einsetzen'

    dfile ='KDPICK.xls';  %"activity.xls" in Einzelzitatmarken ensezten'

lendata   = size(xlsread(sfile,1));
datenzeile = lendata(1);
refn   =  1 ;
desrange =     'B2';

for  i =refn:14
Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));
dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung
xlswrite(dfile,[
xlsread(sfile,1,Range),...
xlsread(sfile,2,Range),...
xlsread(sfile,3,Range),...
xlsread(sfile,4,Range),...
xlsread(sfile,5,Range)], dsheet , desrange);
refn=refn+1;
end
%%

%%%%%%%%%%%%%%%%%%  %%%%%%% Exemple with GRAP                          %%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%
    refz=[    'b'    'c'    'd'    'e'    'f'    'g'    'h'    'i'    'j'    'k'    'l'    'm'    'n'    'o'    ];

    sfile ='KMGRAP.xlsx';  %"KMGRAP.xlsx" in EinzelZitatmarken einsetzen

    dfile ='KDGRAP.xls';  %"activity.xls" in Einzelzitatmarken ensezten

lendata   = size(xlsread(sfile,1));
datenzeile = lendata(1);
refn   =  1 ;
desrange =     'B2';

for  i =refn:14
Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));
dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung
xlswrite(dfile,[
xlsread(sfile,1,Range),...
xlsread(sfile,2,Range),...
xlsread(sfile,3,Range),...
xlsread(sfile,4,Range),...
xlsread(sfile,5,Range),...
xlsread(sfile,6,Range),...
xlsread(sfile,7,Range),...
xlsread(sfile,8,Range)], dsheet , desrange);
refn=refn+1;
end
%%
%%
%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%
refz=[  'b'    'c'    'd'    'e'    'f'    'g'    'h'    'i'    'j'    'k'    'l'    'm'    'n'    'o'    ];

sfile ='KMPLACE.xlsx';  %"KMPLACE.xlsx" in EinzelZitatmarken einsetzen'
dfile ='KDPLACE.xls';  %"activity.xls" in Einzelzitatmarken ensezten'

lendata   = size(xlsread(sfile,1));
datenzeile = lendata(1);
refn   =  1 ;
desrange =     'B2';

for  i =refn:14
Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));
dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung
xlswrite(dfile,[
xlsread(sfile,1,Range),...
xlsread(sfile,2,Range),...
xlsread(sfile,3,Range),...
xlsread(sfile,4,Range),...
xlsread(sfile,5,Range),...
xlsread(sfile,6,Range),...
xlsread(sfile,7,Range),...
xlsread(sfile,8,Range),...
xlsread(sfile,9,Range),...
xlsread(sfile,10,Range)], dsheet , desrange);
refn=refn+1;
end
%%%
%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%

    refz=[    'b'    'c'    'd'    'e'    'f'    'g'    'h'    'i'    'j'    'k'    'l'    'm'    'n'    'o'    ];

    sfile ='KMPUSH.xlsx';  %"KMPUSH.xlsx" in EinzelZitatmarken einsetzen'

    dfile ='KDPUSH.xls';  %"activity.xls" in Einzelzitatmarken ensezten'

lendata   = size(xlsread(sfile,1));
datenzeile = lendata(1);
refn   =  1 ;
desrange =     'B2';

for  i =refn:14
Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));
dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung
xlswrite(dfile,[
xlsread(sfile,1,Range),...
xlsread(sfile,2,Range),...
xlsread(sfile,3,Range),...
xlsread(sfile,4,Range),...
xlsread(sfile,5,Range),...
xlsread(sfile,6,Range)], dsheet , desrange);
refn=refn+1;
end



%%%%%%%%%%%%%%%%%%   Example for MT          %%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 refz=[    'b'    'c'    'd'    'e'    'f'    'g'    'h'    'i'    'j'    'k'    'l'    'm'    'n'    'o'    ];
    sfile ='KMMT.xlsx';  %"KMMT.xlsx" in EinzelZitatmarken einsetzen'

    dfile ='KDMT.xls';  %"activity.xls" in Einzelzitatmarken ensezten'

lendata   = size(xlsread(sfile,1));
datenzeile = lendata(1);
refn   =  1 ;
    desrange =     'B2';

for  i =refn:14
Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));
dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung
xlswrite(dfile,[
xlsread(sfile,1,Range),...
xlsread(sfile,2,Range),...
xlsread(sfile,3,Range),...
xlsread(sfile,4,Range),...
xlsread(sfile,5,Range),...
xlsread(sfile,6,Range)], dsheet , desrange);
refn=refn+1;
end
%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%


%%
%%% Hier kombinieren wir die Probanden für bestimmte Aktivität nach
%%% bestimmte Variable. Die Quelledatei ist die Datei mit Präfix 'KM' und
%%% die Destination Datei ist die Datei mit Präfix 'KD'
%%% Da die Häufigkeit der Aktivität ist nicht alle Prozess nicht gleich
%%% ist,müssen wir zuerst die Maximale Häufigkeit der Aktivität aus
%%% betrachteten Probanden identifizieren( beim lesen SOURCE File der Akivität(Datei mit
%%% Präfix 'KM') [entspricht datenzeile]),
%%% dann alle die Dimension (in Zeile) auf jeder Seite der 'KM'-Datei
%%% uniformisieren. Die Uniformisierung ist besteht hier die leeren Zeile auf
%%% Probendenseite mit Null'0'auffüllen Seite, damit die Datenlänge(Zeile)
%%% auf jede Seite gleich wird.
%%% Mit der Funktion [vorhyooo3] ist der Skript [Meancombinefile] automatisch
%%% generated. With these skript wird einen neuen skript automatisch
%%% generated%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
function[ansheet,nvideo]=vorhyooo3mean( )
% refz={'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'};
% 'Geben Sie die untersuchte Aktivität  als konstante ein!'
dsr= {'desrange = ' 'B2'};
colref = {'refz=[' 'b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o' '];'};
prompt =  {'Geben Sie die untersuchte Aktivität ein!'};
          
dlg_title='The input Variable';
inputd   =inputdlg(prompt,dlg_title);
activity =upper(inputd(1));

adfile = strcat('KM', activity, '.xls');
sfile= strcat('KD',activity, '.xlsx'); % The  excel workbook with activity’s Mean by every Video
nvideo = input('Geben Sie bitte die Anzahl der untersuchten Videos ein:   ');
varz=1;
disp(strcat ('%%%  switch ',' _ ', activity));
disp('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%');
disp(strcat(upper('% Generated Command to copy the  _'), upper('._ Variable')));
% Destination sheet
disp(strcat(' % case ','  ',num2str(varz) ));
disp(upper('%%% Hinweise: Ab hier können sie den generated Code kopieren und löschen und füge die einzelzitatfürhung!!!'))
disp(colref);
disp(strcat('sfile = ',sfile, ';  %"', sfile,'" in EinzelZitatmarken einsetzen'));
disp(strcat('dfile = ',adfile,';  %"activity.xls" in Einzelzitatmarken ensezten'));
disp('lendata   = size(xlsread(sfile,1));');
disp('datenzeile = lendata(1);');
disp('refn   =  1 ;');
disp(dsr);
disp('for  i =refn:14')
disp('Range=strcat(refz(refn),num2str(2),:,refz(refn),num2str(datenzeile));');
disp('dsheet=refn;% The corresponding sheet ist the nummer of the Variable in the source file. Bis hier keine Änderung');
disp(strcat('xlswrite(','dfile',',['));     % The beginning of the Synthax to write the  matrix of the combined sheets
ansheet=1;
    for i=ansheet:nvideo-1		 % Count the sheets in the original Worksbooks
        sheet=num2str(ansheet);
        %The beginning of the Syntax to write the sheets that will be combined. Here we add ',' after 'Range)' to write  the
        %data in  column and '...' to signify to  Matlab to stay on the same line although the command goes to the next line
        disp(strcat('xlsread(','sfile',',',sheet,',','Range',')',',','...'));
        ansheet=ansheet+1;
    end
    
    disp(strcat('xlsread(','sfile',',', num2str(nvideo),',Range',')], dsheet , desrange);'));
    disp('refn=refn+1;');
    disp('end');
    disp('%%%%%%%%%%%%%%%%%%                           %%%%%%%%%%%%');
    disp('% Kopiere die oben angeszeigte Code von "dsheet=... bis dsheet,desrange)" und führen Sie das aus!!')
    disp(upper('% Vielen Dank für ihre Aufmerksamkeit.%'));
    %varz=varz+1;
end 
%%


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%This Function ist to calculate the Anova Fisher for the 14 Variable and
%%%to make the corresponding graphics and make test of Student of the
%%%Medians 
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%         Varianzanalyse Fisher
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function[nvaz act group]= sanomedian2()
close all;clear; 
nvaz=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
%nvazi=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
s_zaehler=1;
act= 'WALK';  group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'	'V111'	'V114'	'V137' 'V141'}; % for WALK
%act= 'PICK';  group=	{'V66'  'V100'	'V111'	'V114'	'V137'	'V141'};      %for PICK	
%act= 'GRAP';  group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'  'V111'};  %for GRAP
%act= 'PLACE'; group=	{'V30'	'V51'	'V56'	'V60'   'V61'  'V64'   'V66'    'V111'	'V114' 'V141'};  %for PLACE
%act= 'PUSH';  group=	{'V51'	'V56'	'V60'	'V61'	'V64'	'V66'};  %for PUSH
%act= 'MT';   group=   {'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'};  %for MT
t= strcat('KMD',act,'.xls');
folder_a=strcat('KRMD',act);
folder_b= strcat('AFM',act);
mkdir(folder_b);
for ij=s_zaehler:14      %%% Da 14 Variable und jede Variable eine Seite
   d = xlsread(t,s_zaehler);%% Offnung der n_te Seite der Datei 
   %dataa = d(:,:);
   [fp,ftable,fstats] = anova1(d,group,'on');%%% Anova Fisher
   title(strcat('Multiple Comparison of mean of Rank of  Variable[',nvaz(s_zaehler),']')); 
   multcompare(fstats,'display','on');%%% Multiple Comparison of statistics aus Fisher ANOVA
   ylabel('Kommissionierer'); title(strcat('Multiple Comparison of mean of Variable[',nvaz(s_zaehler),']')); 
   
   figurelist=findobj('type','figure');
   for j=1:numel(figurelist)
        saveas(figurelist(j),fullfile(folder_b,['figure' num2str(figurelist(j)) '.fig']));% speichert in Format.fig 
        saveas(figurelist(j),fullfile(folder_b,['figure' num2str(figurelist(j)) '.bmp']));  %speichert in Format .bmp 
   end
   s_zaehler=s_zaehler+1;
   
%%%%%% 
end

close(findobj('type','figure'));

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% TEST OF STUDENT%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
cniveau=0.05;
sheet=strcat('NTTEST',act);%%% Sheet for recording t_test value
tfile=strcat('KRMD',act,'.xls'); % This file will contain the t_statistics
ssfile=strcat('MD',act,'.xlsx');
sfs=xlsread(ssfile,act);
[hiii piii ciii]=ttest(sfs,0, cniveau);

xlswrite(tfile,{upper(strcat('t-test for sample analysis _',act))},sheet,'F1');
xlswrite(tfile, nvaz, sheet,'B3:O3');% Datenvariable
xlswrite(tfile, {'HYPOTHESIS'},sheet, 'A4');
xlswrite(tfile,[hiii; piii], sheet, 'B4');
xlswrite(tfile, ciii, sheet, 'B6');% Value von confidence Intervall 
%xlswrite(tfile, nvaz, sheet, 'B6:O6');% Datenvariable for P-Value
xlswrite(tfile,{'PROBABILITY'}, sheet, 'A5');
%xlswrite(tfile,piii, sheet, 'B7:O7'); % p_wert nach test von Student
%st1=upper('Konfidenzintervall mit Signifikanzniveau "');
%stt=cellstr(strcat(st1,num2str(1-cniveau), '"'));
%xlswrite(tfile,stt, sheet, 'F8'); %%% Titel
%xlswrite(tfile, nvaz, sheet, 'B9:O9');% Datenvariable
%xlswrite(tfile, ciii, sheet, 'B10:O11');% Value von confidence Intervall 
xlswrite(tfile, {'LOWER_BOUND';'UPPER_BOUND'}, sheet,'A6:A7');
xlswrite(tfile, {'UNTERSUCHTE VIDEOS FÜR AKTIVITÄT_'}, sheet,'F9');
xlswrite(tfile, {act}, sheet, 'I9');
xlswrite(tfile, group, sheet, 'A10');


copyfile(folder_b,folder_a );%%% Copy the folder_b in folder_a
copyfile(tfile,folder_a)%%% Copy the Result file ('KRMD',act,'.xls') in folder_a
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% ZIELE DIESER FUNKTION %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%% Anwendung von zwei Methoden für die Varianzanalyse der gesamten Daten jeder Aktivität und nach jeden
%%% Variablen: Ncht parametrische Test(KRUSKALWALLIS TEST ANOVA ONE WAY),
%%% weil alle Daten sind nicht normalverteilt UND DANN mit standard
%%% ANOVA(ANOVA mit p_wert von FISHER) basierend auf die Approximation der Daten zur
%%% Normalverteilung mit dem zentralen Grenzwertsatz                                                 
%%% Ergebnisse Zusammenfassen und Vergleich von der Ergebnisse der 2
%%% Methoden%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function[nva,table, p stats ]= fanotest3med()
%%% Here we performe ANOVA KRUSKALWALLIS; ANOVA FISHER WITHOUT GRAPH AND RECORD THE RESULT IN FILE 
clc;clear;close all; 
nva=[{'AccX[m/s^2]'} {'AccY[m/s^2]'} {'AccZ[m/s^2]'} {'GyrX[degree/sec]'} {'GyrY[degree/sec]'} {'GyrZ[degree/sec]'}... 
     {'MagX[uT]'} {'MagY[uT]'} {'MagZ[uT]'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
 
nvaz=[{'AccX'} {'AccY'} {'AccZ'} {'GyrX'} {'GyrY'} {'GyrZ'} {'MagX'} {'MagY'} {'MagZ'} {'Q0'} {'Q1'} {'Q2'} {'Q3'} {'Vbat'}];
%nvazm=['AccX' 'AccY' 'AccZ' 'GyrX' 'GyrY' 'GyrZ' 'MagX' 'MagY' 'MagZ' 'Q0' 'Q1' 'Q2' 'Q3' 'Vbat'];

refz=['b' 'c' 'd' 'e' 'f' 'g' 'h' 'i' 'j' 'k' 'l' 'm' 'n' 'o'];
s_zaehler = 1;

act= 'WALK';  group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'	'V111'	'V114'	'V137' 'V141'}; % for WALK
%act= 'PICK';  group=	{'V66'  'V100'	'V111'	'V114'	'V137'	'V141'};      %for PICK	
%act= 'GRAP';  group=	{'V30'	'V51'	'V56'	'V60'	'V61'	'V64'  'V111'};  %for GRAP
%act= 'PLACE'; group=	{'V30'	'V51'	'V56'	'V60'   'V61'  'V64'   'V66'    'V111'	'V114' 'V141'};  %for PLACE
%act= 'PUSH';  group=	{'V51'	'V56'	'V60'	'V61'	'V64'	'V66'};  %for PUSH
%act= 'MT';   group=   {'V51'	'V56'	'V60'	'V61'	'V64'	'V66'	'V100'};  %for MT

t=strcat('KMD',upper(act),'.xls'); %t=uigetfile;
dimension=size(xlsread(t,1));%Dateidimension inklusive Beschriftung(Probanden oder Group in ersten Zeile)
m=dimension(1);   % minus 1 wenn die gruppe Zahlen sind weil die Zeile der Gruppe in Datei soll nicht in berechnung angewendet werden 
n=dimension(2);   % oder n=size(group); % Anzahl der Probanden. 



desfile=strcat('KRMD',act,'.xls');
% Der Name der gewählten Datei oben
folder_a=strcat('KRMD',act);
mkdir(folder_a);
folder_b= strcat('AFM',act);
mkdir(folder_b);

for i=s_zaehler:14      %%% Da 14 Variable und jede Variable eine Seite
   d = xlsread(t,s_zaehler);%% Offnung der n_te Seite der Datei 
   datatest = d(1:m,1:n);
   
   [p,table,stats] = kruskalwallis(datatest,group,'on');%%% Nicht-parametrische Test  Anova
   
   b1=strcat('Varianzanalyse/',act,'/', nvaz(s_zaehler),'/Kruskal-Wallis One-Way ANOVA/', 'PWert=', num2str(p));
   b2=strcat('Untersuchung der Aktivität _', act, '_ bei den Probanden an der Variablen: ',nva(s_zaehler));
   b3={'Varianzanalyse:Kruskal-Wallis One-Way ANOVA TABLE/Variable'};
   k1={b2};
   kb1=cellstr(b1);
   kb2=cellstr(b2); 
   title(b1);% Beschriftung der ANOVA Graphik
   ylabel(strcat('Ranks of mean of[',nvaz(s_zaehler),']'));
   
   [fp,ftable] = anova1(datatest,group,'on');%%% Anova Fisher
   disp(b2); disp(b3); disp(table);
   xlswrite(desfile,b1, s_zaehler,'b1');
   xlswrite(desfile,kb2, s_zaehler,'b2');
   xlswrite(desfile,b3, s_zaehler,'b3'); 
   xlswrite(desfile,nva(s_zaehler), s_zaehler,'G3'); % Untersuchte Variable
   
   %title(strcat('Varianzanalyse für**',act,': Kruskal-Wallis One-Way ANOVA :  ', 'p-Wert=', num2str(p)))
  
   title(strcat(act,'/ANOVA ONE WAY/Variable', nva(s_zaehler)));
   xlabel('Kommissionierer');
   ylabel(nva(s_zaehler));
   disp(strcat('Untersuchung der**', act, '** bei den Kommissionierer:',group, 'an der Variablen: ',nva(s_zaehler)));
   disp('Varianzanalyse: Kruskal-Wallis  One-Way ANOVA');

   if p>0.05
       decision=strcat('Da PWert ',' = ',num2str(p),'   >   0.05;_');
       b4=strcat('** ENTSCHEIDUNG für** ', nva(s_zaehler));
       b5= ' Hypothese H1 verwerfen! Differenz zwischen Probanden NICHT SIGNIFIKANT';
       disp(b4); disp(b5);
       b6= strcat(decision,'  ',b4);
       b5= {b5};
       b6= cell(b6); 
       
       xlswrite(desfile,b6 ,s_zaehler,'b9');
       xlswrite(desfile,b5 ,s_zaehler,'b10');
       
       xlabel('Kommissionierer! Differenz nicht signifikant');
   else
      decision=strcat('Da PWert ',' = ',num2str(p),'   <  0.05 ;');
      b4=strcat('** ENTSCHEIDUNG für** ', nva(s_zaehler));
      b5='Hypothese H0 verwerfen! Differenz zwischen Probanden SIGNIFIKANT. ';
      disp(b4); disp(b5);
      b6= strcat(decision,'  ',b4);
      b5= {b5};
      b6= cell(b6);
      xlabel('Probanden/Kommissionier! Differenz zwischen Probanden ist signifikant.');
      
      xlswrite(desfile,b6,s_zaehler ,'b9');
      xlswrite(desfile,b5 ,s_zaehler,'b10');
   end
   
   %%%%% Multiple Comparison of mean of Variable
   
   multcompare(stats,'display','on');%%% Multiple Comparison of statistics aus Kruskalwallis ANOVA
   ylabel('Kommissionierer');
   title(strcat('Multiple Comparison of mean ranks of Variable[',nvaz(s_zaehler),']'));
   
   %%%% Test Statistik Kruskalwallis in Excelblatt schreiben 
   
   xlswrite(desfile,table, s_zaehler,'b4');%Kopie von Anova Table
   xlswrite(desfile,{'Untersuchte Videos'},s_zaehler,'d21');
   xlswrite(desfile, group,s_zaehler,'b22');% Untersuchte Gruppe

   %%%% Test Statistik Fisher Anova in Excelblatt schreiben
   xlswrite(desfile, {'ANOVA TABLE FISHER'},s_zaehler, 'd12');
   xlswrite(desfile,ftable, s_zaehler,'b13');%Kopie von Anova Table
   %xlswrite(desfile,fstatistics, s_zaehler,'b28');%%% Kopie von Statististics von ANOVA 
   %save(tfile, nva(s_zaehler));
   %save(tfile, fstats);
   if p>0.05
       df=strcat('PWert ',' =',num2str(fp), '>0.05', '_Hypothese H1 verwerfen!');
       decision={df};
       xlswrite(desfile,decision,s_zaehler,'B18');
       xlswrite(desfile,{'Differenz zwischen Kommissionierern ist NICHT SIGNIFIKANT'},s_zaehler,'B19');
   else
       df=strcat('PWert ',' =',num2str(fp), '<0.05', '_Hypothese H0 verwerfen!');
       decision={df};
       xlswrite(desfile,decision,s_zaehler,'B18')
       xlswrite(desfile,{'Differenz zwischen Kommissionierern ist SIGNIFIKANT'},s_zaehler,'B19'); 
   end

   
   %%%%%%% GRAPHIK SPEICHERN
   figurelist=findobj('type','figure');
   for j=1:numel(figurelist)
        saveas(figurelist(j),fullfile(folder_a,['figure' num2str(figurelist(j)) '.fig']));% speichert in Format.fig 
        saveas(figurelist(j),fullfile(folder_a,['figure' num2str(figurelist(j)) '.bmp']));  %speichert in Format .bmp 
   end
   
   s_zaehler=s_zaehler+1;
end
close(findobj('type','figure'));
   
%%%% ZUSAMMENFASSUNG DER VARIANZANALYSE FÜR DIE AKTIVITÄT 

zusf=strcat('AKTIVITÄT :', act,':ZUSAMMENFASSUNG DER VARIANZANALYSE');
zusammenfassung={zusf};
tsheet='zusammenfassung';
xlswrite(desfile,zusammenfassung,tsheet, 'F2');%%% Überschrift der Seite
xlswrite(desfile,nvaz,tsheet, 'B3');%%% Die Variablen
xlswrite(desfile,{'VARIABLE'},tsheet,'A3');
xlswrite(desfile,{'p_Wert   - aus KRUSKALWALLIS ANOVA'},tsheet,'A4');
xlswrite(desfile,{'p_Wert   - aus FISHER ANOVA'},tsheet,'A6');
xlswrite(desfile,{'HYPOTHESE  *** Kruskalwallis'},tsheet,'A5');
xlswrite(desfile,{'HYPOTHESE - ANOVA- Fisher '},tsheet,'A7');
xlswrite(desfile,{'HYPOTHESE - ANOVA- Fisher '},tsheet,'A7');
xlswrite(desfile,{'UNTERSUCHTE VIDEOS'},tsheet,'E12');
xlswrite(desfile,group,tsheet,'B13');
xlswrite(desfile,{'HYPOTHESE= 1; bedeutet H0 verwerfen! dh.**Differenz zwischen Kommissionierern ist  SIGNIFIKANT  '},tsheet,'B9');
xlswrite(desfile,{'HYPOTHESE= 0; bedeutet H1 verwerfen! dh.**Differenz zwischen Kommissionierern ist  NICHT SIGNIFIKANT'},tsheet,'B10');
w=1;
for x=w:14
    refzk='G5';
    refzf='G14';
    %%% Bildung der Referenze  oder Zellenposition 
    drefzk= strcat(refz(w),num2str(4));
    drefzf= strcat(refz(w),num2str(6));
    hrefzk= strcat(refz(w),num2str(5));
    hrefzf= strcat(refz(w),num2str(7));
    %%% P_Wert der Varianzanalyse lesen und kopieren in der Seite der
    %%% Zusamenfassung
    p_wk=xlsread(desfile,w,refzk);
    p_wf=xlsread(desfile,w,refzf); 
    xlswrite(desfile,p_wk,tsheet,drefzk);
    xlswrite(desfile,p_wf,tsheet,drefzf);
    
    %%% Welche Hypothese besagt den  p_Wert und ihre Bedeutung ? 
    if p_wk > 0.05
       xlswrite(desfile,'0',tsheet,hrefzk); %%% H0 annehmen
    else
       xlswrite(desfile,'1',tsheet,hrefzk); %%% H1 annehmen
    end
    if p_wf > 0.05
       xlswrite(desfile,'0',tsheet,hrefzf); %%% H0 annehmen: Differenz nicht Signifikant
    else
       xlswrite(desfile,'1',tsheet,hrefzf); %%% H1 annehmen: Differenz Signifikant
    end
    w=w+1; 
    
end
copyfile(desfile,folder_a);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%This Function ist to calculate the Anova Fisher for the 14 Variable and
%%%to make the corresponding graphics and make test of Student of the
%%%Medians 
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%         Varianzanalyse Fisher
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close all; 


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% TEST OF STUDENT%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
cniveau=0.05;
sheet=strcat('NTTEST',act);%%% Sheet for recording t_test value
tfile=strcat('KRMD',act,'.xls'); % This file will contain the t_statistics
ssfile=strcat('MD',act,'.xlsx');
sfs=xlsread(ssfile,act);
[hiii piii ciii]=ttest(sfs,0, cniveau);

xlswrite(tfile,{upper(strcat('t-test for sample analysis _',act))},sheet,'F1');
xlswrite(tfile, nvaz, sheet,'B3:O3');% Datenvariable
xlswrite(tfile, {'HYPOTHESIS'},sheet, 'A4');
xlswrite(tfile,[hiii; piii], sheet, 'B4');
xlswrite(tfile, ciii, sheet, 'B6');% Value von confidence Intervall 
%xlswrite(tfile, nvaz, sheet, 'B6:O6');% Datenvariable for P-Value
xlswrite(tfile,{'PROBABILITY'}, sheet, 'A5');
xlswrite(tfile, {'LOWER_BOUND';'UPPER_BOUND'}, sheet,'A6:A7');
xlswrite(tfile, {'UNTERSUCHTE VIDEOS FÜR AKTIVITÄT_'}, sheet,'F9');
xlswrite(tfile, {act}, sheet, 'I9');
xlswrite(tfile, group, sheet, 'A10');


copyfile(folder_b,folder_a );%%% Copy the folder_b in folder_a
copyfile(tfile,folder_a)%%% Copy the Result file ('KRMD',act,'.xls') in folder_a
clear;