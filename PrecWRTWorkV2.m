%Calculate rate of change of precision for change in unit of work
clear;
clc;
close all;
%Try 50000, should take 10 hours. Maybe check 100000 time on HP. HP looks
%like it will take about 12 hours.
its=100000;
bigfx=24;
medfx=17;
smallfx=10;
fileflag=1;
diaryname=strcat('WorkSimOpt_tab0_its',int2str(its),'.txt')
diary(diaryname)
diary on

alltime=cputime;

%params of form [Mx,Mx,max x to count in linear method transect, max number of full count FOVs,omega]
params=[30000,30000,1100,75,2];
params=[params;30000,10000,1100,75,2];
params=[params;30000,5000,1100,120,2];
params=[params;30000,3000,1100,150,2];
params=[params;30000,2000,1100,200,2];
params=[params;30000,500,1100,350,2];
nosets=size(params,1)

for i=1:nosets
%for i=1:1
    Mx=params(i,1);
    y3bar=Mx*(3*3)/(100*100);
    Mn=params(i,2);
    uhat=Mx/Mn;
    maxtlim=params(i,3);
%     fx=params(i,4);
%     if fx>21
%         error('XXXXXXXXXXX fx must be <= 21 XXXXXXXXXXX')
%     end
    maxfn=params(i,4);
    omega=params(i,5);
    deltastari=uhat*sqrt((omega+y3bar)/(omega*uhat+y3bar));
    optimalfnbig=ceil(deltastari*bigfx);
    optimalfnmed=ceil(deltastari*medfx);
    optimalfnsmall=ceil(deltastari*smallfx);

    fprintf('param set = %i, y3bar = %f, uhat = %d, omega = %d, deltastari = %f\n', i, y3bar, uhat, omega,deltastari)
    fprintf('fx = %i, fn = %f,fx = %i, fn = %f,fx = %i, fn = %f\n',bigfx,optimalfnbig,medfx,optimalfnmed,smallfx,optimalfnsmall)
    %fprintf('fx = %i, fn = %f,fx = %i, fn = %f,fx = %i, fn = %f\n',bigfx,ceil(1.25*deltastari*bigfx),medfx,ceil(1.25*deltastari*medfx),smallfx,ceil(1.25*deltastari*smallfx))
    
    %maxfnbig=min(max(ceil(1.25*optimalfnbig),50),19*19-bigfx)
    maxfnbig=19*19-bigfx
    fprintf(2,'[fx=%i] batch i=%d, Mx=%d, Mn=%d, maxtlim=%d, maxfn=%d, %s\n',bigfx,i,Mx,Mn,maxtlim,maxfnbig,datestr(clock));
    %WorkSimV2(Mx,Mn,maxtlim,bigfx,maxfn,its,fileflag);
    [LPrecWorkOpt,FPrecWorkOptBig]=WorkSimV2(Mx,Mn,maxtlim,bigfx,maxfnbig,its,fileflag);
    
    %maxfnmed=min(max(ceil(1.25*optimalfnmed),50),19*19-medfx)
    maxfnmed=19*19-medfx
    fprintf(2,'[fx=%i] batch i=%d, Mx=%d, Mn=%d, maxtlim=%d, maxfn=%d, %s\n',medfx,i,Mx,Mn,maxtlim,maxfnmed,datestr(clock));
    [LPrecWorkOpt,FPrecWorkOptMed]=WorkSimV2(Mx,Mn,maxtlim,medfx,maxfnmed,its,fileflag);
    
    %maxfnsmall=min(max(ceil(1.25*optimalfnsmall),50),19*19-smallfx)
    maxfnsmall=19*19-smallfx
    fprintf(2,'[fx=%i] batch i=%d, Mx=%d, Mn=%d, maxtlim=%d, maxfn=%d, %s\n',smallfx,i,Mx,Mn,maxtlim,maxfnsmall,datestr(clock));
    [LPrecWorkOpt,FPrecWorkOptSmall]=WorkSimV2(Mx,Mn,maxtlim,smallfx,maxfnsmall,its,fileflag);
    

    %LPrecWorkOpt of form: [no. fossils in transect; no. markers in transect; effort; precision; concentration estimate; % error in concentration estimate; effort beyond which accuracy goal is met]
    %FPrecWorkOpt of form: [no. calibration FOVs; no. full count FOVs; effort; precision;no. fossils in cal count (fixed);no. exotics in full count; concentration estimate; % error in concentration estimate; effort beyond which accuracy goal is met]
    
    ConcLstartwork=LPrecWorkOpt(end:end);
    if ConcLstartwork==-1
        fprintf(2,'Linear method not met accuracy goal\n');
    end
    ConcFstartworkBig=FPrecWorkOptBig(end:end);
    if ConcFstartworkBig==-1
        fprintf(2,'FOVS method (BIG) not met accuracy goal\n');
    end
    ConcFstartworkMed=FPrecWorkOptMed(end:end);
    if ConcFstartworkMed==-1
        fprintf(2,'FOVS method (MED) not met accuracy goal\n');
    end
    ConcFstartworkSmall=FPrecWorkOptSmall(end:end);
    if ConcFstartworkSmall==-1
        fprintf(2,'FOVS method (SMALL) not met accuracy goal\n');
    end
    
    
    mw1=max(LPrecWorkOpt(3,:));
    mw2=max(FPrecWorkOptBig(3,:));
    mw3=max(FPrecWorkOptMed(3,:));
    mw4=max(FPrecWorkOptSmall(3,:));
    maxwork=max([mw1,mw2,mw3,mw4]);
    
    mp1=max(LPrecWorkOpt(4,:));
    mp2=max(FPrecWorkOptBig(4,:));
    mp3=max(FPrecWorkOptMed(4,:));
    mp4=max(FPrecWorkOptSmall(4,:));
    maxprec=max([mp1,mp2,mp3,mp4]);
    
    %cla
    figure(i)
    hold on
    scatter(FPrecWorkOptBig(3,:),FPrecWorkOptBig(4,:),'k','+');
    scatter(FPrecWorkOptMed(3,:),FPrecWorkOptMed(4,:),'b','*');
    scatter(FPrecWorkOptSmall(3,:),FPrecWorkOptSmall(4,:),'r','x');
    % xline(ConcFstartworkBig,'--k');
    % xline(ConcFstartworkMed,'--b');
    % xline(ConcFstartworkSmall,'--r');
    if optimalfnbig<=length(FPrecWorkOptBig)
        optimallinebig=FPrecWorkOptBig(3,optimalfnbig);
    else
        optimallinebig=0;
    end
    fprintf('For fx = %i, optimal fn = %i, optimal effort = %f\n',bigfx, optimalfnbig,optimallinebig)
    if optimalfnmed<=length(FPrecWorkOptMed)
        optimallinemed=FPrecWorkOptMed(3,optimalfnmed);
    else
        optimallinemed=0;
    end
    fprintf('For fx = %i, optimal fn = %i, optimal effort = %f\n',medfx, optimalfnmed,optimallinemed)
    if optimalfnsmall<=length(FPrecWorkOptSmall)
        optimallinesmall=FPrecWorkOptSmall(3,optimalfnsmall);
    else
        optimallinesmall=0;
    end
    fprintf('For fx = %i, optimal fn = %i, optimal effort = %f\n',smallfx, optimalfnsmall,optimallinesmall)
    
    xline(optimallinebig,'--k');%%%Need to fix these so taht they correspond to the correct fn value
    xline(optimallinemed,'--b');
    xline(optimallinesmall,'--r');
    scatter(LPrecWorkOpt(3,:),LPrecWorkOpt(4,:),[],[0.9290 0.6940 0.1250],'.');
    %xline(ConcLstartwork,'-.','Color',[0.9290 0.6940 0.1250]);
    xlim([0,1.1*maxwork])
    ylim([0,1.1*maxprec])
    title(['Prec/Effort, ratio=',int2str(Mx/Mn)])
    hold off
    fprintf('XXXXXXXXXXXXXXXXXXX\nXXXXXXXXXXXXXXXXXXX\n\n')
end

fprintf('total time= %f, at %s\n',cputime-alltime,datestr(clock))