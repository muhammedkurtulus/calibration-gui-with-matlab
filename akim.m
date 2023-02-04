function varargout = akim(varargin)
% EXCEL MATLAB code for excel.fig
%      EXCEL, by itself, creates a new EXCEL or raises the existing
%      singleton*.
%
%      H = EXCEL returns the handle to a new EXCEL or the handle to
%      the existing singleton*.
%
%      EXCEL('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in EXCEL.M with the given input arguments.
%
%      EXCEL('Property','Value',...) creates a new EXCEL or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before akim_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to akim_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help excel

% Last Modified by GUIDE v2.5 06-Oct-2020 15:51:49

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @akim_OpeningFcn, ...
                   'gui_OutputFcn',  @akim_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before excel is made visible.
function akim_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to excel (see VARARGIN)

% Choose default command line output for excel
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes excel wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = akim_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;
% global obj1 obj2;
% % Find a VISA-GPIB object.
% obj1 = instrfind('Type', 'visa-gpib', 'RsrcName', 'GPIB0::1::INSTR', 'Tag', '');
% obj2 = instrfind('Type', 'visa-gpib', 'RsrcName', 'GPIB0::18::INSTR', 'Tag', '');
% % Read serial port objects from memory to MATLAB workspace
% % Create the VISA-GPIB object if it does not exist
% % otherwise use the object that was found.
% if isempty(obj1)
%     obj1 = visa('NÝ', 'GPIB0::1::INSTR');
% else
%     fclose(obj1);
%     obj1 = obj1(1);
% end
% if isempty(obj2)
%     obj2 = visa('NÝ', 'GPIB0::18::INSTR');
% else
%     fclose(obj2);
%     obj2 = obj2(1);
% end
% % Connect to instrument object, obj1.
% obj1.InputBufferSize = 5100; %buffer boyutu default 512 bytes
% fopen(obj1);
% fopen(obj2);
%To change the default for that you need to change Preferences -> Variables
%format long;

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%global obj1 deger sapma ort;
global obj1 obj2 a b ort1 ort2 ort3 ssapma1 ssapma2 ssapma3 gedeger1 gedeger2 gedeger3 hata1 hata2 hata3

set(handles.text2,'string','Ýþlem sürdürülüyor...');
pause(0.1); %bildirim gözükmesi için
nom=get(handles.edit1,'string');
load('akimdeger.mat');

% nomcommand = ['SOUR:PHAS:CURR:MHAR:HARM1 '  nom];
% fprintf(obj2,'UNIT:MHAR:CURR ABS');%Uygulanacak olan gerilimin 1. harmoniðinin (sinüs) mutlak deðeriyle  çalýþýlacaðýný gösteren komuttur
% fprintf(obj2,'SOUR:PHAS:CURR:RANG 0.5,5');%Uygulanacak olan gerilimin hangi çalýþma bölgesinde olacaðýný tanýmlamak için kullanýlýr. Örneðin, bu bölge min. 23 V max. 336 V RMS uygulanabileceðini gösterir.
% %fprintf(obj2,'SOUR:PHAS:VOLT:MHAR:HARM1 230,0');%Uygulanacak olan gerilimin RMS deðeri ve fazýný belirtmek için kullanýlýr. Faz burada 0 derece
% fprintf(obj2,nomcommand);
% fprintf(obj2,'SOUR:FREQ 50');%frekans giriþi 
% fprintf(obj2,'SOUR:PHAS:CURR:STAT 1');%Girilen fonksiyonun(yardim) pasif ya da aktif olmasýný saðlar. 1 ON demektir ve gerilimin uygulamaya hazýr hale getirilmesini belirtir. 
% fprintf(obj2,'OUTP:STAT ON');%Cihazýn girilen fonksiyonlarý uygulamasýný ya da ksmesini saðlar. ON-OFF
% pause(2);
% fprintf(obj1,'RESET');
% fprintf(obj1,'TARM HOLD');% Okuduðunu tut.(ekrana gelen deðeri hafýzada tut)
% fprintf(obj1,'ACV');% AC yardim modunu seç.
% fprintf(obj1,'TRIG AUTO');% Her ölçümü kendi zamanlamasý ile yapmasýný saðlar. Örneðin AUTO yerine SYNC deseydik þebeke frekansýyla senkron olarak ölçüm alacaktý.
% fprintf(obj1,'NRDGS 10, AUTO');% her tetiklemede kaç ölçüm alýnacaðý hakkýnda oto örnekleme ayarý
% fprintf(obj1,'MEM FIFO');% hafýzayý aç
% fprintf(obj1,'TARM SGL, 3');% kaç kez tetikleme gerçekleþtirileceði bilgisi
% pause(0.1);
% deger = fscanf(obj1);
% fprintf(obj2,'OUTP:STAT OFF');

set(handles.text2,'string',' ');
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
ss = strsplit(deger,' '); %gelen veriler ayrýþtýrýlýr.
newStr=ss(2:end); %1. sütunda kalan boþ karakter diziden çýkarýlýr. 
%set(handles.listbox2,'string',deger);
%c={deger};
%newStr(29)=newStr(28);
newStr(30)=newStr(27);
new1=newStr(1:10);
new2=newStr(11:20);
new3=newStr(21:30);
a = cell(10,3); %10 satýr 1 sütun
for i=1:10
    
    a{i,1}=str2double(new1(i));
    a{i,2}=str2double(new2(i));
    a{i,3}=str2double(new3(i));
end


%set(handles.uitable4,'Data',a,'ColumnName','2'); %Data
b = cell(10,3);
for i=1:10
    
    b{i,1}=num2str(a{i,1},6);
    b{i,2}=num2str(a{i,2},6);
    b{i,3}=num2str(a{i,3},6);
end
set(handles.uitable2,'Data',b);
%%%%%%%%%%%%%%ortalama%%%%%%%%%%%%%%%%%%%%%

%std=std(newStr);
toplam1=0;
toplam2=0;
toplam3=0;

%display(newStr);
for i=1:10
    toplam1=str2double(new1(i))+toplam1;
    toplam2=str2double(new2(i))+toplam2;
    toplam3=str2double(new3(i))+toplam3;
end
ort1=toplam1/10;
ort2=toplam2/10;
ort3=toplam3/10;
%ort=mean(a);


sum1=0;
sum2=0;
sum3=0;
%%%%%%%%%%%%% varyans & standart sapma %%%%%%%%%%%%%%%%%%
for i=1:10
    x1=(str2double(new1(i))-ort1)^2;
    x2=(str2double(new2(i))-ort2)^2;
    x3=(str2double(new3(i))-ort3)^2;
    sum1=x1+sum1;
    sum2=x2+sum2;
    sum3=x3+sum3;
end
varyans1=sum1/(10-1);
varyans2=sum2/(10-1);
varyans3=sum3/(10-1);
ssapma1=sqrt(varyans1);
ssapma2=sqrt(varyans2);
ssapma3=sqrt(varyans3);
cal=xlsread('/Kalibrasyon.xls','Sayfa1','B2');
gedeger1=ort1*(1/cal);
gedeger2=ort2*(1/cal);
gedeger3=ort3*(1/cal);

nom2 = strsplit(nom,','); 
nom3=nom2(1);

hata1=(str2double(nom3)-gedeger1)/str2double(nom3)*100;
hata2=(str2double(nom3)-gedeger2)/str2double(nom3)*100;
hata3=(str2double(nom3)-gedeger3)/str2double(nom3)*100;
%disp(new3);
set(handles.text7,'string',num2str(ort1,6)); %0 hariç basamak sayýsý
set(handles.text15,'string',num2str(ort2,6));
set(handles.text17,'string',num2str(ort3,6));
set(handles.text8,'string',num2str(ssapma1,2));
set(handles.text16,'string',num2str(ssapma2,2));
set(handles.text22,'string',num2str(ssapma3,2));
set(handles.text14,'string',num2str(cal,7));
set(handles.text10,'string',num2str(gedeger1,7));
set(handles.text20,'string',num2str(gedeger2,7));
set(handles.text21,'string',num2str(gedeger3,7));
set(handles.text12,'string',num2str(hata1,2));
set(handles.text18,'string',num2str(hata2,2));
set(handles.text19,'string',num2str(hata3,2));

msgbox('Tamamlandý');
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
closereq();

% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global b
c = cell(10,3);
for i=1:10
    c{i,1}=str2double(b{i,1});
    c{i,2}=str2double(b{i,2});
    c{i,3}=str2double(b{i,3});
end

xlswrite('data.xls',c,'Akim','B2');
xlswrite('data.xls',str2double(get(handles.text7,'string')),'Akim','B12');
xlswrite('data.xls',str2double(get(handles.text15,'string')),'Akim','C12');
xlswrite('data.xls',str2double(get(handles.text17,'string')),'Akim','D12');
xlswrite('data.xls',str2double(get(handles.text8,'string')),'Akim','B15');
xlswrite('data.xls',str2double(get(handles.text16,'string')),'Akim','C15');
xlswrite('data.xls',str2double(get(handles.text22,'string')),'Akim','D15');
xlswrite('data.xls',str2double(get(handles.text10,'string')),'Akim','B13');
xlswrite('data.xls',str2double(get(handles.text20,'string')),'Akim','C13');
xlswrite('data.xls',str2double(get(handles.text21,'string')),'Akim','D13');
xlswrite('data.xls',str2double(get(handles.text12,'string')),'Akim','B14');
xlswrite('data.xls',str2double(get(handles.text18,'string')),'Akim','C14');
xlswrite('data.xls',str2double(get(handles.text19,'string')),'Akim','D14');

pause(0.1);
winopen('data.xls');

% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
winopen('Sertifika.doc');


% --------------------------------------------------------------------
function gerilim_Callback(hObject, eventdata, handles)
% hObject    handle to gerilim (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
uiwait(msgbox('Gerilim kalibrasyonu için baðlantýlarýnýzý yapýnýz'));
gerilim; %akým kalibrasyonu uygulamasý açýlýr.

% --------------------------------------------------------------------
function excel_Callback(hObject, eventdata, handles)
% hObject    handle to Untitled_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
answer = questdlg('Dikkat! Bütün Excel iþlemleri sonlandýrýlacaktýr, onaylamadan önce lütfen iþlemlerinizi kaydediniz.', ...
	'Uyarý', ...
	'Devam et','Vazgec','Vazgec');
% Handle response
switch answer
    case 'Devam et'
        system('taskkill /F /IM EXCEL.EXE'); %varsa excel iþlemlerini kapatýr.
        msgbox('Tamamlandý. Kaydetmek için tekrardan "Kaydet" butonuna basýnýz.')
    case 'Vazgec'
        return;
end
