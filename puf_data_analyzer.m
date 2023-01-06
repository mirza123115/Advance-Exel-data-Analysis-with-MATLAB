
clear all 
close all
clc

workbook='D:\WF_ana\TILT_analysis_wFB_100mV _exp_WL0.7.xlsx';

t=0:7;



for j=1:12;
[vt_x,dlt_x,dlc_x] = importfile(workbook,j,3,114);
VT_y(:,j)=vt_x;
dlt_y(:,j)=dlt_x;
dlc_y(:,j)=dlc_x;


VT_adder=VT_y(:,j);
dlt=dlt_y(:,j);
dlc=dlc_y(:,j);



%[VT_adder,dlt,dlc] = importfile('TILT_analysis_wFB_100mV _exp_WL0.8.xlsx','tt25',3,114);
% data=xlsread('TILT_analysis_wFB_100mV _exp_WL0.8.xlsx','tt25','E3:E114');
% data(:,2:3)=xlsread('TILT_analysis_wFB_100mV _exp_WL0.8.xlsx','tt25','H3:I114');
% VT_adder=data(:,1);
% dlt=data(:,2);
% dlc=data(:,3);


%index initialization
y=0;
z=0;

for i=0:7;
    
    if i==0;
        y=1;
        z=14;
        
    else
   
        y=14*i+1;
        z=14*i+14;
    end
    
% figure(i+1)

dlt1=dlt(y:z);
dlc1=dlc(y:z);
vt1=VT_adder(y:z);
dif_x=(dlt1-dlc1);
dif=abs(dlt1-dlc1); %difference of dlt and dlc
a=min(dif);  %finding closer point to the intersect point
b=find(a==dif); %index of the min difference value
% c(i+1)=(b+14*i); %index off all vt_adder within whole VT_adder (1-112)
% d(i+1)=VT_adder(c(i+1)); %value of vt adder closing to the intersecting point by calling their intersecting point

if (b>1) && (abs(dif_x(b)+dif_x(b-1))<abs(dif_x(b))+abs(dif_x(b-1)))
    b=b-1;
end

       
%straight line equation-------------
% syms x
% y_dlt= dlt1(b) + ((x-vt1(b))/(vt1(b)-vt1(b+1)))*(dlt1(b)-dlt1(b+1)); %dlt vs vt adder
% y_dlc= dlc1(b) + ((x-vt1(b))/(vt1(b)-vt1(b+1)))*(dlc1(b)-dlc1(b+1)); %dlc vs vt adder
% vt(i+1)=solve(y_dlt-y_dlc,0)

if b<14
    y_dlt= @(x) dlt1(b) + ((x-vt1(b))/(vt1(b)-vt1(b+1)))*(dlt1(b)-dlt1(b+1)); %dlt vs vt adder
    y_dlc= @(x) dlc1(b) + ((x-vt1(b))/(vt1(b)-vt1(b+1)))*(dlc1(b)-dlc1(b+1)); %dlc vs vt adder
    p= @(x) (dlt1(b) + ((x-vt1(b))/(vt1(b)-vt1(b+1)))*(dlt1(b)-dlt1(b+1)))- (dlc1(b) + ((x-vt1(b))/(vt1(b)-vt1(b+1)))*(dlc1(b)-dlc1(b+1)));
    q=fsolve(p,0);
    
    vt(i+1,j)=q;
else
    vt(i+1,j)=300;
end
    

if q<0
    vt(i+1,j)=0;
end
    

%plot(VT_adder(y:z),dlt(y:z),VT_adder(y:z),dlc(y:z));
%plot(VT_adder(y:z),dlt1,VT_adder(y:z),dlc1);
% plot(vt1,dlt1,vt1,dlc1);

end
plot(t,vt(:,j))
hold on
end

vt=round(vt,2);

xlswrite(workbook,vt,13,'B12:M19')

%difference between manual and automation
xlswrite(workbook,vt-xlsread(workbook,13,'B2:M9'),13,'B22:M29')


% Clear temporary variables
clearvars VT_adder vt_x vt1 dlt dlt_x dlt1 dlc dlc_x dlc1 i j  a b p q t y z ;

