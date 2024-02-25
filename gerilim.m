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
    obj1 = visa('N�', 'GPIB0::1::INSTR');
else
    fclose(obj1);
    obj1 = obj1(1);
end
if isempty(obj2)
    obj2 = visa('N�', 'GPIB0::18::INSTR');
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

set(handles.text3,'string','��lem s�rd�r�l�yor...');
pause(0.1); %bildirim g�z�kmesi i�in

nom=get(handles.nominal,'string');
nomcommand = ['SOUR:PHAS:VOLT:MHAR:HARM1 '  nom];
fprintf(obj2,'UNIT:MHAR:VOLT ABS');%Uygulanacak olan gerilimin 1. harmoni�inin (sin�s) mutlak de�eriyle  �al���laca��n� g�steren komuttur
fprintf(obj2,'SOUR:PHAS:VOLT:RANG 23,336');%Uygulanacak olan gerilimin hangi �al��ma b�lgesinde olaca��n� tan�mlamak i�in kullan�l�r. �rne�in, bu b�lge min. 23 V max. 336 V RMS uygulanabilece�ini g�sterir.
%fprintf(obj2,'SOUR:PHAS:VOLT:MHAR:HARM1 230,0');%Uygulanacak olan gerilimin RMS de�eri ve faz�n� belirtmek i�in kullan�l�r. Faz burada 0 derece
fprintf(obj2,nomcommand);
fprintf(obj2,'SOUR:FREQ 50');%frekans giri�i 
fprintf(obj2,'SOUR:PHAS:VOLT:STAT 1');%Girilen fonksiyonun(gerilim) pasif ya da aktif olmas�n� sa�lar. 1 ON demektir ve gerilimin uygulamaya haz�r hale getirilmesini belirtir. 
fprintf(obj2,'OUTP:STAT ON');%Cihaz�n girilen fonksiyonlar� uygulamas�n� ya da ksmesini sa�lar. ON-OFF

pause(2);

fprintf(obj1,'RESET');
fprintf(obj1,'TARM HOLD');% Okudu�unu tut.(ekrana gelen de�eri haf�zada tut)
fprintf(obj1,'ACV');% AC Gerilim modunu se�.
fprintf(obj1,'TRIG AUTO');% Her �l��m� kendi zamanlamas� ile yapmas�n� sa�lar. �rne�in AUTO yerine SYNC deseydik �ebeke frekans�yla senkron olarak �l��m alacakt�.
fprintf(obj1,'NRDGS 10, AUTO');% her tetiklemede ka� �l��m al�naca�� hakk�nda oto �rnekleme ayar�
fprintf(obj1,'MEM FIFO');% haf�zay� a�
fprintf(obj1,'TARM SGL, 3');% ka� kez tetikleme ger�ekle�tirilece�i bilgisi

deger = fscanf(obj1);
%pause(0.1);

fprintf(obj2,'OUTP:STAT OFF');

set(handles.text3,'string',' ');

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
ss = strsplit(deger,' '); %gelen veriler ayr��t�r�l�r.
newStr=ss(2:end); %1. s�tunda kalan bo� karakter diziden ��kar�l�r. 
newStr(30)=newStr(27); %30.veri gelmedi�i i�in 27. de�ere e�itledik

new1=newStr(1:10);
new2=newStr(11:20);
new3=newStr(21:30);

a = cell(10,3); %10 sat�r 1 s�tun
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
%set(handles.tablo,'Data',a,'ColumnName','2'); %istenilen s�tuna yazd�r�labilir

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

set(handles.ortalama1,'string',num2str(ort1,6)); %0 hari� basamak say�s�
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

answer = questdlg('Gerilim kalibrasyonu tamamlanm��t�r. Ak�m kalibrasyonu i�in ba�lant�lar�n�z� yap�n�z.', ...
	'Uyar�', ...
	'Devam et','Kapat','Kapat');
% Handle response
switch answer
    case 'Devam et'
        akim; %ak�m kalibrasyonu uygulamas� a��l�r.
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

% --- Word'den okuma kodlar�------------------
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

answer = questdlg('Dikkat! B�t�n Excel i�lemleri sonland�r�lacakt�r, onaylamadan �nce l�tfen i�lemlerinizi kaydediniz.', ...
	'Uyar�', ...
	'Devam et','Vazgec','Vazgec');
% Handle response
switch answer
    case 'Devam et'
        system('taskkill /F /IM EXCEL.EXE'); %varsa excel i�lemlerini sonland�r�r.
        msgbox('Tamamland�. Kaydetmek i�in tekrardan "Kaydet" butonuna bas�n�z.')
    case 'Vazgec'
        return;
end


% --------------------------------------------------------------------
function akim_Callback(hObject, eventdata, handles)
% hObject    handle to akim (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

uiwait(msgbox('Ak�m kalibrasyonu i�in ba�lant�lar�n�z� yap�n�z'));
akim; %ak�m kalibrasyonu uygulamas� a��l�r.
