function varargout = D_123190015_responsiSAW(varargin)
% D_123190015_RESPONSISAW MATLAB code for D_123190015_responsiSAW.fig
%      D_123190015_RESPONSISAW, by itself, creates a new D_123190015_RESPONSISAW or raises the existing
%      singleton*.
%
%      H = D_123190015_RESPONSISAW returns the handle to a new D_123190015_RESPONSISAW or the handle to
%      the existing singleton*.
%
%      D_123190015_RESPONSISAW('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in D_123190015_RESPONSISAW.M with the given input arguments.
%
%      D_123190015_RESPONSISAW('Property','Value',...) creates a new D_123190015_RESPONSISAW or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before D_123190015_responsiSAW_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to D_123190015_responsiSAW_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help D_123190015_responsiSAW

% Last Modified by GUIDE v2.5 26-Jun-2021 00:04:11

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @D_123190015_responsiSAW_OpeningFcn, ...
                   'gui_OutputFcn',  @D_123190015_responsiSAW_OutputFcn, ...
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


% --- Executes just before D_123190015_responsiSAW is made visible.
function D_123190015_responsiSAW_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to D_123190015_responsiSAW (see VARARGIN)

% Choose default command line output for D_123190015_responsiSAW
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes D_123190015_responsiSAW wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = D_123190015_responsiSAW_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in tampildata.
function tampildata_Callback(hObject, eventdata, handles)
% hObject    handle to tampildata (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%data = xlsread('DATA RUMAH.xlsx','C2:H21');
%set(handles.table1,'data',data);    
opts = spreadsheetImportOptions("NumVariables", 8);

% Menentukan sheet yag digunakan dan data mana
% saja yang akan digunakan
opts.Sheet = "Sheet1";
opts.DataRange = "A2:H21";

% Menentukan nama variabel serta tipe data apa saja yang akan digunakan
opts.VariableNames = ["No", "Var2", "HR", "LB", "LT", "JKT", "JKM", "JGars"];
opts.SelectedVariableNames = ["No", "HR", "LB", "LT", "JKT", "JKM", "JGars"];
opts.VariableTypes = ["double", "char", "double", "double", "double","double", "double", "double"];

opts = setvaropts(opts, "Var2", "WhitespaceRule", "preserve");
opts = setvaropts(opts, "Var2", "EmptyFieldRule", "auto");

% Mengambil data dari file excel
data = readmatrix("DATA RUMAH.xlsx", opts);
%tampilkan data di tabel
set(handles.table1,'data',data);

% --- Executes on button press in prosesdata.
function prosesdata_Callback(hObject, eventdata, handles)
% hObject    handle to prosesdata (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%Nilai bobot tiap kriteria
w = [0.3,0.2,0.23,0.1,0.07,0.1];

%Atribut tiap-tiap kriteria, dimana nilai 1 = benefit, dan 0 = cost
k = [0,1,1,1,1,1];

%load excel
a = xlsread('DATA RUMAH.xlsx','C2:H1011');

%Ambil jumlah baris dan kolom dari data
[f,g]=size (a);

R=zeros (f,g); %Buat matriks kosong R
Y=zeros (f,g); %Buat matriks kosong Y

for j=1:g
    if k(j)==1  %kriteria dengan atribut benefit
        R(:,j)=a(:,j)./max(a(:,j));
    else %kriteria dengan atribut cost
        R(:,j)=min(a(:,j))./a(:,j);
    end
end

%Mulai proses perankingan
for i=1:f
    V(i)= sum(w.*R(i,:));
end
ranking = sort(V, 'descend');
B = ranking.';

%cari nilai terbesar dan indexnya
[nilaiterbesar, letak] = max(V);
%cetak hasil di command window
disp("No Rumah : "+letak +", Letak : " +nilaiterbesar);

%cetak 20 data terbaik di tabel
set(handles.table3, 'data', B(1:20,:)); 
