clear;
clc;
[Supply1_data,Supply1_name]=xlsread("Supply2010_2014.csv");
[Supply2_data,Supply2_name]=xlsread("Supply2015_2019.csv");
Supply1_name(1,:)=[];
Supply2_name(1,:)=[];
Supply_alldata=[Supply1_data;Supply2_data];
Supply_allname=[Supply1_name;Supply2_name];
[Feed_data,Feed_name]=xlsread("Feed.xlsx"); %Feed
Feed_name(1,:)=[];
[Export_data,Export_name]=xlsread("Export.xlsx"); %Export
Export_name(1,:)=[];
[Import_data,Import_name]=xlsread("Import.xlsx"); %import
Import_name(1,:)=[];
[Losses_data,Losses_name]=xlsread("Losses.xlsx"); %Losses
Losses_name(1,:)=[];
[OtherUse_data,OtherUse_name]=xlsread("OtherUse.xlsx"); %OtherUse
OtherUse_name(1,:)=[];
[Production_data,Production_name]=xlsread("Production.xlsx"); %Production
Production_name(1,:)=[];
[Seed_data,Seed_name]=xlsread("Seed.xlsx"); %Seed
Seed_name(1,:)=[];
[StockVariation_data,StockVariation_name]=xlsread("StockVariation.xlsx"); %StockVariation
StockVariation_name(1,:)=[];
%%筛选年份
year=2010;
A=find(Supply_alldata(:,7)==year);%Supply
Supply_alldata=Supply_alldata(A,:);
Supply_allname=Supply_allname(A,:);
B=find(Feed_data(:,7)==year);%Feed
Feed_data=Feed_data(B,:);
Feed_name=Feed_name(B,:);
C=find(Export_data(:,7)==year);%Export
Export_data=Export_data(C,:);
Export_name=Export_name(C,:);
D=find(Import_data(:,7)==year);%Import
Import_data=Import_data(D,:);
Import_name=Import_name(D,:);
E=find(Losses_data(:,7)==year);%Losses
Losses_data=Losses_data(E,:);
Losses_name=Losses_name(E,:);
F=find(OtherUse_data(:,7)==year);%OtherUse
OtherUse_data=OtherUse_data(F,:);
OtherUse_name=OtherUse_name(F,:);
G=find(Production_data(:,7)==year);%Production
Production_data=Production_data(G,:);
Production_name=Production_name(G,:);
H=find(Seed_data(:,7)==year);%Seed
Seed_data=Seed_data(H,:);
Seed_name=Seed_name(H,:);
I=find(StockVariation_data(:,7)==year);%StockVariation
StockVariation_data=StockVariation_data(I,:);
StockVariation_name=StockVariation_name(I,:);
%%合并
[col_Supply,row_Supply]=size(Supply_alldata);
result=zeros(col_Supply,9);
result(:,1)=Supply_alldata(:,10); %第一列放Supply
%Feed筛选
tic;
[col_Feed,row_Feed]=size(Feed_name);
for i=1:col_Feed
    AA=strmatch(Feed_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(Feed_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,2)=Feed_data(i,10); %第二列放Feed
end
toc;
%Export筛选
tic;
[col_Export,row_Export]=size(Export_name);
for i=1:col_Export
    AA=strmatch(Export_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(Export_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,3)=Export_data(i,10);  %第三列放Export
end
toc;
%Import筛选
tic;
[col_Import,row_Import]=size(Import_name);
for i=1:col_Import
    AA=strmatch(Import_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(Import_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,4)=Import_data(i,10);  %第四列放Import
end
toc;
%Losses筛选
tic;
[col_Losses,row_Losses]=size(Losses_name);
for i=1:col_Losses
    AA=strmatch(Losses_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(Losses_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,5)=Losses_data(i,10);  %第五列放Import
end
toc;
%OtherUse筛选
tic;
[col_OtherUse,row_OtherUse]=size(OtherUse_name);
for i=1:col_OtherUse
    AA=strmatch(OtherUse_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(OtherUse_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,6)=OtherUse_data(i,10);  %第六列放OtherUse
end
toc;
%Production筛选
tic;
[col_Production,row_Production]=size(Production_name);
for i=1:col_Production
    AA=strmatch(Production_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(Production_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,7)=Production_data(i,10);  %第七列放Production
end
toc;
%Seed筛选
tic;
[col_Seed,row_Seed]=size(Seed_name);
for i=1:col_Seed
    AA=strmatch(Seed_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(Seed_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,8)=Seed_data(i,10);  %第八列放Seed
end
toc;
%StockVariation筛选
tic;
[col_StockVariation,row_StockVariation]=size(StockVariation_name);
for i=1:col_StockVariation
    AA=strmatch(StockVariation_name(i,4),Supply_allname(:,4),'exact'); %找到出口国的行号(这里需要循环）
    AAA=strmatch(StockVariation_name(i,8),Supply_allname(:,8),'exact');
    AAAAA=intersect(AA,AAA);
    result(AAAAA,9)=StockVariation_data(i,10);  %第九列放StockVariation
end
toc;
xlswrite(strcat('E:\Data0827\',num2str(year),'_data.xlsx'),result);
xlswrite(strcat('E:\Data0827\',num2str(year),'_name.xlsx'),Supply_allname);
