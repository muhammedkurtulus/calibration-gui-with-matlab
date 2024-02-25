function varargout = gerilim(varargin)
% GERILIM MATLAB code for gerilim.fig
%      GERILIM, by itself, creates a new GERILIM or raises the existing
%      singleton*.
%
%      H = GERILIM returns the handle to a new GERILIM or the handle to
%      the existing singleton*.
%
%      GERILIM('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GERILIM.M with the given input arguments.
%
%      GERILIM('Property','Value',...) creates a new GERILIM or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before gerilim_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to gerilim_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help gerilim

% Last Modified by GUIDE v2.5 06-Oct-2020 16:03:44

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @gerilim_OpeningFcn, ...
                   'gui_OutputFcn',  @gerilim_OutputFcn, ...
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


% --- Executes just before gerilim is made visible.
function gerilim_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to gerilim (see VARARGIN)

% Choose default command line output for gerilim
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes gerilim wait for user response (see UIRESUME)
% uiwait(handles.figure1);

% --- Outputs from this function are returned to the command line.
function varargout = gerilim_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

global obj1 obj2;
 
% Find a VISA-GPIB object.

obj1 = instrfind('Type', 'visa-gpib', 'RsrcName', 'GPIB0::1::INSTR', 'Tag', '');
obj2 = instrfind('Type', 'visa-gpib', 'RsrcName', 'GPIB0::18::INSTR', 'Tag', '');

% Read serial port objects from memory to MATLAB workspace
% Create the VISA-GPIB object if it does not exist
% otherwise use the object that was found.

if isempty(obj1)
    obj1 = visa('NÝ', 'GPIB0::1::INSTR');
else
    fclose(obj1);
    obj1 = obj1(1);
end
if isempty(obj2)
    obj2 = visa('NÝ', 'GPIB0::18::INSTR');
else
    fclose(obj2);
    obj2 = obj2(1);
end

obj1.InputBufferSize = 5100; %buffer boyutu default 512 bytes

% Connect to instrument object.
fopen(obj1);
fopen(obj2);

% --- Executes on button press in baslat.
function baslat_Callback(hObject, eventdata, handles)
% hObject    handle to baslat (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

global obj1 obj2 b ort1 ort2 ort3 ssapma1 ssapma2 ssapma3 gedeger1 gedeger2 gedeger3 hata1 hata2 hata3

set(handles.text3,'string','Ýþlem sürdürülüyor...');
pause(0.1); %bildirim gözükmesi için

nom=get(handles.nominal,'string');
nomcommand = ['SOUR:PHAS:VOLT:MHAR:HARM1 '  nom];
fprintf(obj2,'UNIT:MHAR:VOLT ABS');%Uygulanacak olan gerilimin 1. harmoniðinin (sinüs) mutlak deðeriyle  çalýþýlacaðýný gösteren komuttur
fprintf(obj2,'SOUR:PHAS:VOLT:RANG 23,336');%Uygulanacak olan gerilimin hangi çalýþma bölgesinde olacaðýný tanýmlamak için kullanýlýr. Örneðin, bu bölge min. 23 V max. 336 V RMS uygulanabileceðini gösterir.
%fprintf(obj2,'SOUR:PHAS:VOLT:MHAR:HARM1 230,0');%Uygulanacak olan gerilimin RMS deðeri ve fazýný belirtmek için kullanýlýr. Faz burada 0 derece
fprintf(obj2,nomcommand);
fprintf(obj2,'SOUR:FREQ 50');%frekans giriþi 
fprintf(obj2,'SOUR:PHAS:VOLT:STAT 1');%Girilen fonksiyonun(gerilim) pasif ya da aktif olmasýný saðlar. 1 ON demektir ve gerilimin uygulamaya hazýr hale getirilmesini belirtir. 
fprintf(obj2,'OUTP:STAT ON');%Cihazýn girilen fonksiyonlarý uygulamasýný ya da ksmesini saðlar. ON-OFF

pause(2);

fprintf(obj1,'RESET');
fprintf(obj1,'TARM HOLD');% Okuduðunu tut.(ekrana gelen deðeri hafýzada tut)
fprintf(obj1,'ACV');% AC Gerilim modunu seç.
fprintf(obj1,'TRIG AUTO');% Her ölçümü kendi zamanlamasý ile yapmasýný saðlar. Örneðin AUTO yerine SYNC deseydik þebeke frekansýyla senkron olarak ölçüm alacaktý.
fprintf(obj1,'NRDGS 10, AUTO');% her tetiklemede kaç ölçüm alýnacaðý hakkýnda oto örnekleme ayarý
fprintf(obj1,'MEM FIFO');% hafýzayý aç
fprintf(obj1,'TARM SGL, 3');% kaç kez tetikleme gerçekleþtirileceði bilgisi

deger = fscanf(obj1);
%pause(0.1);

fprintf(obj2,'OUTP:STAT OFF');

set(handles.text3,'string',' ');

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
ss = strsplit(deger,' '); %gelen veriler ayrýþtýrýlýr.
newStr=ss(2:end); %1. sütunda kalan boþ karakter diziden çýkarýlýr. 
newStr(30)=newStr(27); %30.veri gelmediði için 27. deðere eþitledik

new1=newStr(1:10);
new2=newStr(11:20);
new3=newStr(21:30);

a = cell(10,3); %10 satýr 1 sütun
for i=1:10
    a{i,1}=str2double(new1(i));
    a{i,2}=str2double(new2(i));
    a{i,3}=str2double(new3(i));
end


b = cell(10,3);
for i=1:10
    b{i,1}=num2str(a{i,1},6);
    b{i,2}=num2str(a{i,2},6);
    b{i,3}=num2str(a{i,3},6);
end

set(handles.tablo,'Data',b);
%set(handles.tablo,'Data',a,'ColumnName','2'); %istenilen sütuna yazdýrýlabilir

%%%%%%%%%%%%%%ortalama%%%%%%%%%%%%%%%%%%%%%
toplam1=0;
toplam2=0;
toplam3=0;

for i=1:10
    toplam1=str2double(new1(i))+toplam1;
    toplam2=str2double(new2(i))+toplam2;
    toplam3=str2double(new3(i))+toplam3;
end

ort1=toplam1/10;
ort2=toplam2/10;
ort3=toplam3/10;

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

cal=xlsread('/Kalibrasyon.xls','Sayfa1','A2');

gedeger1=ort1*cal;
gedeger2=ort2*cal;
gedeger3=ort3*cal;

nom2 = strsplit(nom,','); 
nom3=nom2(1);

hata1=(str2double(nom3)-gedeger1)/str2double(nom3)*100;
hata2=(str2double(nom3)-gedeger2)/str2double(nom3)*100;
hata3=(str2double(nom3)-gedeger3)/str2double(nom3)*100;

set(handles.ortalama1,'string',num2str(ort1,6)); %0 hariç basamak sayýsý
set(handles.ortalama2,'string',num2str(ort2,6));
set(handles.ortalama3,'string',num2str(ort3,6));
set(handles.ssapma1,'string',num2str(ssapma1,2));
set(handles.ssapma2,'string',num2str(ssapma2,2));
set(handles.ssapma3,'string',num2str(ssapma3,2));
set(handles.kalibrasyon,'string',num2str(cal,7));
set(handles.gedeger1,'string',num2str(gedeger1,7));
set(handles.gedeger2,'string',num2str(gedeger2,7));
set(handles.gedeger3,'string',num2str(gedeger3,7));
set(handles.hata1,'string',num2str(hata1,2));
set(handles.hata2,'string',num2str(hata2,2));
set(handles.hata3,'string',num2str(hata3,2));

answer = questdlg('Gerilim kalibrasyonu tamamlanmýþtýr. Akým kalibrasyonu için baðlantýlarýnýzý yapýnýz.', ...
	'Uyarý', ...
	'Devam et','Kapat','Kapat');
% Handle response
switch answer
    case 'Devam et'
        akim; %akým kalibrasyonu uygulamasý açýlýr.
    case 'Kapat'
        return;
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% --- Executes on button press in kapat.
function kapat_Callback(hObject, eventdata, handles)
% hObject    handle to kapat (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% global obj1 obj2
% fclose(obj1);
% fclose(obj2);
%close(handles.figure1);
closereq(); 

% --- Executes on button press in kaydet.
function kaydet_Callback(hObject, eventdata, handles)
% hObject    handle to kaydet (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

global b

c = cell(10,3);
for i=1:10
    c{i,1}=str2double(b{i,1});
    c{i,2}=str2double(b{i,2});
    c{i,3}=str2double(b{i,3});
end

xlswrite('data.xls',c,'Gerilim','B2');
xlswrite('data.xls',str2double(get(handles.ortalama1,'string')),'Gerilim','B12');
xlswrite('data.xls',str2double(get(handles.ortalama2,'string')),'Gerilim','C12');
xlswrite('data.xls',str2double(get(handles.ortalama3,'string')),'Gerilim','D12');
xlswrite('data.xls',str2double(get(handles.ssapma1,'string')),'Gerilim','B15');
xlswrite('data.xls',str2double(get(handles.ssapma2,'string')),'Gerilim','C15');
xlswrite('data.xls',str2double(get(handles.ssapma3,'string')),'Gerilim','D15');
xlswrite('data.xls',str2double(get(handles.gedeger1,'string')),'Gerilim','B13');
xlswrite('data.xls',str2double(get(handles.gedeger2,'string')),'Gerilim','C13');
xlswrite('data.xls',str2double(get(handles.gedeger3,'string')),'Gerilim','D13');
xlswrite('data.xls',str2double(get(handles.hata1,'string')),'Gerilim','B14');
xlswrite('data.xls',str2double(get(handles.hata2,'string')),'Gerilim','C14');
xlswrite('data.xls',str2double(get(handles.hata3,'string')),'Gerilim','D14');

winopen('data.xls');

% --- Word'den okuma kodlarý------------------
% hObject    handle to pushbutton4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% global sometext;
% word = actxserver('Word.Application');
% wdoc = word.Documents.Open('C:\Users\muham\Desktop\Staj\kalb.docx');
% %wdoc is the Document object which you can query and navigate.
% sometext = wdoc.Content.Text;
% set(handles.text15,'string',sometext);
% wdoc.Close; % close document
% word.Quit;  % end application
%---------------------------------------------


% --- Executes on button press in sertifika.
function sertifika_Callback(hObject, eventdata, handles)
% hObject    handle to sertifika (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

winopen('Sertifika.doc');


% --- Executes on mouse press over figure background, over a disabled or
% --- inactive control, or over an axes background.


% --------------------------------------------------------------------
function excel_Callback(hObject, eventdata, handles)
% hObject    handle to excel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

answer = questdlg('Dikkat! Bütün Excel iþlemleri sonlandýrýlacaktýr, onaylamadan önce lütfen iþlemlerinizi kaydediniz.', ...
	'Uyarý', ...
	'Devam et','Vazgec','Vazgec');
% Handle response
switch answer
    case 'Devam et'
        system('taskkill /F /IM EXCEL.EXE'); %varsa excel iþlemlerini sonlandýrýr.
        msgbox('Tamamlandý. Kaydetmek için tekrardan "Kaydet" butonuna basýnýz.')
    case 'Vazgec'
        return;
end


% --------------------------------------------------------------------
function akim_Callback(hObject, eventdata, handles)
% hObject    handle to akim (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

uiwait(msgbox('Akým kalibrasyonu için baðlantýlarýnýzý yapýnýz'));
akim; %akým kalibrasyonu uygulamasý açýlýr.
