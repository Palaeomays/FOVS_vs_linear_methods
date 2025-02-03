clc;
alltime=cputime;
its=1000000;
tmaster=0;%Tablet error
fileflag=1;
diaryname=strcat('BigFossilSimsV3_Opt_T',int2str(10^4*tmaster),'_its',int2str(its),'.txt')
diary(diaryname)
diary on

%high density
%Mxmaster=60000;

%medium density
Mxmaster=30000; %Gives y3bar=27. y3bar average is 26 in Lyco productivity data - eastern Australia basins10 (export-22-8-2023).xlsx

%low density
%Mxmaster=15000; %15000 gives smallest density with this simulation number of FOVs for effort = 1000


%params of form: [no. fossils, no. markers, tablet error contribution]
params=[Mxmaster,Mxmaster,tmaster;Mxmaster,round(Mxmaster*(2.5)/3),tmaster;Mxmaster,round(Mxmaster*2/3),tmaster;Mxmaster,round(Mxmaster/2),tmaster;Mxmaster,round(Mxmaster/3),tmaster;Mxmaster,round(Mxmaster/6),tmaster;Mxmaster,round(Mxmaster/10),tmaster;Mxmaster,round(Mxmaster/15),tmaster;Mxmaster,round(Mxmaster/20),tmaster;Mxmaster,round(Mxmaster/30),tmaster;Mxmaster,round(Mxmaster/60),tmaster]


%For calculating required number of FOVs and number of linear method x to
%count
omega=2;
%choose fixed effort or fixed error > 0
error=-1;%!!FIXED ERROR NOT WORKING YET!!
    if error>-1
        error('XXX code for specified error not finished yet XXX\n')
    end
effort=1000;

nosets=size(params,1)
ratios=(params(:,1)./params(:,2))';
tlims=zeros(1,nosets);
fxs=zeros(1,nosets);
fns=zeros(1,nosets);
Lerrors=zeros(1,nosets);
LsimEs=zeros(1,nosets);
LtrueEs=zeros(1,nosets);
Lefforts=zeros(1,nosets);
Ferrors=zeros(1,nosets);
FsimEs=zeros(1,nosets);
FtrueEs=zeros(1,nosets);
Fefforts=zeros(1,nosets);
pvalues=zeros(1,nosets);

fprintf('XXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n\n\n')

for i=1:nosets
    fprintf('batch i=%d, started at %s\n',i,datestr(clock));
    Mx=params(i,1);
    Mn=params(i,2);
    t=params(i,3);
    rho=Mx/Mn;
    y3bar=Mx*(3*3)/(100*100);
    a=(omega/y3bar)+1+(1/rho);
    
    %linear method parameter
    tlim=round(effort/a);
    tlims(i)=tlim;
    
    %FOVS method parameters
    [N3star,fstar,FOVratio,FOVxdensity]=FOVoptimiserV1(Mx,Mn,omega,t,error,effort);
    N3star=round(N3star);
    fstar=round(fstar);
    fxs(i)=N3star;
    fns(i)=fstar;
    
    if (effort<0)&&(error>0)
        fprintf('Aiming for error = %f%%\n',error)
    end
    
    if (effort>0)&&(error<0)
        fprintf('Aiming for effort = %f units\n',effort)
    end
    opty3num=rho^2+sqrt(rho^3*(1+rho*(rho-1)));
    opty3den=(rho+1)*(rho-1)^2;
    y3barstar=2*omega*opty3num/opty3den;
    
    if y3bar>=y3barstar
        fprintf(2,'FOVS preferred: rho= %f, y3bar= %f, y3barstar= %f\n',rho,y3bar,y3barstar);
    end
    if y3bar<y3barstar
        fprintf(2,'LINEAR preferred: rho= %f, y3bar= %f, y3barstar= %f\n',rho,y3bar,y3barstar);
    end
        
    fprintf('Mx=%d, Mn=%d, t=%f, tlim=%d, fx=%d, fn=%d\n',Mx,Mn,t,tlim,N3star,fstar);
    
    
    [a1,a2,a3,a4,a5,a6,a7,a8,a9]=MicrofossilSimV3(Mx,Mn,t,tlim,N3star,fstar,its,fileflag);
    Lerrors(i)=a1;
    LsimEs(i)=a2;
    LtrueEs(i)=a3;
    Lefforts(i)=a4;
    Ferrors(i)=a5;
    FsimEs(i)=a6;
    FtrueEs(i)=a7;
    Fefforts(i)=a8;
    pvalues(i)=a9;
end

% figure(10)
% scatter(ratios,Lerrors,90,'.','b')
% title('Standard errors of linear (blue) and FOV (red)')
% hold on
% scatter(ratios,Ferrors,40,'+','r')
% hold off
% 
% figure(11)
% scatter(ratios,LtrueEs,90,'.','b')
% title('True errors of linear (blue) and FOV (red)')
% hold on
% scatter(ratios,FtrueEs,40,'+','r')
% hold off
% 
% figure(12)
% scatter(ratios,Lefforts,90,'.','b')
% title('Efforts of linear (blue) and FOV (red)')
% hold on
% scatter(ratios,Fefforts,40,'+','r')
% hold off

%Simulation output data of form:
%[Mx,Mn,t,tlim,fx,fn,linear total error,linear sim std dev, linear true std dev,linear efforts,FOV total error,FOV sim std dev,FOV true std dev,FOV efforts, p values for Wilcoxon test between total errors]
opt=[params,tlims',fxs',fns',Lerrors',LsimEs',LtrueEs',Lefforts',Ferrors',FsimEs',FtrueEs',Fefforts',pvalues']

if fileflag==1
    filenom10=strcat('.\SimData\FossilSims_T',int2str(10^4*tmaster),'_its',int2str(its),'.csv')
    writematrix(opt,filenom10);
end

fprintf('total time = %f, at %s\n',cputime-alltime,datestr(clock))

diary off