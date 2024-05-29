function [caltotalx,calmeanx,ftotalx,fmeanx,ftotaln,fmeann,Fconc,xhats] = SimStatsChecker(Mx,Mn,tlim,fn,its,fopt)
diaryname=strcat('PoissonMarkerCheck_its',int2str(its),'.txt')
diary(diaryname)
diary on
fprintf(2,'Mx= %d, Mn= %d, fn= %d, its= %d\nStarted at %s\n',Mx,Mn,fn,its,datestr(clock))
noprints=10;%The number of print statements to produce while the code is running to record the time taken
format compact;
format long;

%building the calibration fields of view
fcal=15; %Number of calibration fields of view
aperturewidth=3;
apertureheight=3;
leftsx=[4.5:11:99,48.5*ones(1,6)];
bottomsx=[48.5*ones(1,9),10,23,36,61,74,87];

rightsx=leftsx+aperturewidth;
topsx=bottomsx+apertureheight;
fovsx=[leftsx;rightsx;bottomsx;topsx];%fields of view with format [left edge; right edge; bottom edge; top edge]

%building the full count fields of view
block=floor(fn/50);%First 50 FOVs in left block, then next 50 FOVs in right block
nonews=mod(fn,50);%Number of FOVs in the new block
newlefts=zeros(1,nonews);%List of left edges of FOVs in new block
newbottoms=zeros(1,nonews);%List of bottom edges of FOVs in new block

for fj=1:nonews
     row=floor((fj-1)/5);
     if fj<=25
         botj=2+10*row;
     else
         botj=5+10*row;
     end
     leftj=2+10*mod(fj-1,5);
     if block==1
         leftj=leftj+53;
     end
     newlefts(fj)=leftj;
     newbottoms(fj)=botj;
end

if block==1
    leftsn=[repmat(2:10:42,1,10),newlefts];
    bottomsn=[2*ones(1,5),12*ones(1,5),22*ones(1,5),32*ones(1,5),42*ones(1,5),55*ones(1,5),65*ones(1,5),75*ones(1,5),85*ones(1,5),95*ones(1,5)];
    bottomsn=[bottomsn,newbottoms];
else
    leftsn=newlefts;
    bottomsn=newbottoms;
end

rightsn=leftsn+aperturewidth;
topsn=bottomsn+apertureheight;
fovsn=[leftsn;rightsn;bottomsn;topsn];%fields of view with format [left edge; right edge; bottom edge; top edge]

ttotalx=zeros(its,1); %total x in transect (linear)
ttotaln=zeros(its,1); %total n in transect (linear)
tsdx=zeros(its,1); %st dev of x in transect, sqrt(x)/meanx(x) (linear)
tsdn=zeros(its,1); %st dev of n in transect, sqrt(n)/meanx(n) (linear)
transrights=zeros(its,1); %Number of steps needed to find transect. Only used to evaluate efficiency of program.

caltotalx=zeros(its,1); %calibration count of x in all FOVs
calmeanx=zeros(its,1); %calibration mean number of x per FOV
calvarx=zeros(its,1); %calibration variance of mean number of x
calsdx=zeros(its,1); %calibration st dev of mean number of x

ftotalx=zeros(its,1); %full count of x in all FOVs (explicit counting, as opposed to estimating using mean from calibration)
fmeanx=zeros(its,1); %mean number of x in full count (explicit mean, as opposed to using mean from calibration)
fsdx=zeros(its,1); %st dev of mean number of x in full count (explicit st dev, , as opposed to using st dev from calibration
ftotaln=zeros(its,1); %count of exotics in full count in all FOVs
fmeann=zeros(its,1); %mean number of exotics per FOV
fsdn=zeros(its,1); %st dev of mean number of exotics per FOV

xhats=zeros(its,1);%Estimated number of total fossils in full count (using calibration count)
Fconc=zeros(its,1);%Concentration calculation for FOV method (assuming V=1)
tic;
%treturn of form [transect x count,std dev of x count;transect n count, std dev of n count]
%calreturn of form [counted x in all calibration FOVs,mean x in calibration FOVs, std dev of counted x in calibration FOVs]
%fullreturn of form:
    %row 1= [counted x in all full FOVs, mean x in full FOVs, std dev of x in full FOVs, estimated total x]
    %row 2= [counted n in all full FOVs, mean n in full FOVs, std dev of n in full FOVs, concentration]
fnnobins=20;
fnhist=zeros(1,fnnobins);

for i=1:its
    if i==its
        [treturn,calreturn,fullreturn,newfnhist]=MicrofossilSim_iCheck(Mx,Mn,tlim,fovsx,fovsn,1,fnnobins);
    else
        [treturn,calreturn,fullreturn,newfnhist]=MicrofossilSim_iCheck(Mx,Mn,tlim,fovsx,fovsn,0,fnnobins);
    end
    fnhist=fnhist+newfnhist;
    
    transrights(i)=treturn(1,3); %Right edge of transect
    ttotalx(i)=treturn(1,1);%total x in transect (linear)
    ttotaln(i)=treturn(2,1);%total n in transect (linear)
    tsdx(i)=treturn(1,2);%st dev of x in transect, sqrt(x)/meanx(x) (linear)
    tsdn(i)=treturn(2,2);%st dev of n in transect, sqrt(n)/meanx(n) (linear)
    
    caltotalx(i)=calreturn(1); %calibration count of x in all FOVs
    calmeanx(i)=calreturn(2); %calibration mean number of x per FOV
    calvarx(i)=calreturn(3);
    calsdx(i)=calreturn(4); %calibration st dev of mean number of x (proportional, corrected with c4)

    ftotalx(i)=fullreturn(1,1); %full count of x in all FOVs (explicit counting, as opposed to estimating using mean from calibration)
    fmeanx(i)=fullreturn(1,2); %mean number of x in full count (explicit mean, as opposed to using mean from calibration)
    fsdx(i)=fullreturn(1,3); %st dev of mean number of x in full count (explicit st dev, as opposed to using st dev from calibration
    ftotaln(i)=fullreturn(2,1); %count of exotics in full count in all FOVs
    fmeann(i)=fullreturn(2,2); %mean number of exotics per FOV
    fsdn(i)=fullreturn(2,3); %st dev of mean number of exotics in full count
    
    xhats(i)=fullreturn(1,4); %Estimated number of total fossils in full count (using calibration count)
    Fconc(i)=fullreturn(2,4); %Concentration calculation (assuming V=1)
    %fprintf('time for i = %d is %f secs at %s\n',i,toc,datestr(clock));
    if mod(i,floor(its/noprints))==0
        fprintf('i= %d, at %s\n',i,datestr(clock))
    end
end

%Single FOV
prob=(3*3)/(100*100);
fnhistN=fnhist/sum(fnhist);
lambdahat=mean(fmeann)
mean(ftotaln)/fn
fprintf('per FOV: exact lambda= %f, est lambda= %f\n',prob*Mn,lambdahat)
xs=0:(fnnobins-1);
y1=poisspdf(xs,lambdahat);
binoprobs=binopdf(xs,Mn,prob);
SEregP1=sum((y1-fnhistN).^2)
SEregB1=sum((binoprobs-fnhistN).^2)


figure(20)
hold off
bar(xs,fnhistN)
title('histogram of marker counts per FOV, Poisson [red], Binomial [blue]')
xlim([-0.5,9])
hold on
plot(xs,y1,'r','LineWidth',5)
plot(xs,binoprobs,'b')
hold off

figure(21)
hold off
bar(xs,fnhistN)
title('histogram of marker counts per FOV, Poisson [red]')
xlim([-0.5,9])
hold on
plot(xs,y1,'r','LineWidth',3)
hold off

%Sum of all FOVs
totallambdahat=mean(ftotaln);
fprintf('all FOVs: exact lambda= %f, est lambda= %f\n',prob*Mn*fn,totallambdahat)
nototalbins=50;
totalxs=0:nototalbins-1;
totaly1=poisspdf(totalxs,totallambdahat);
totalbinoprobs=binopdf(totalxs,Mn,fn*prob);

histftotaln=histcounts(ftotaln,0:nototalbins,'Normalization','pdf');

SEregP2=sum((totaly1-histftotaln).^2)
SEregB2=sum((totalbinoprobs-histftotaln).^2)


% figure(21)
% histogram(ftotaln,'Normalization','pdf')
% xlim([-1,30])
% title('histogram of marker counts in all full count FOVs, Poisson [red], Binomial [blue]')
% hold on
% plot(totalxs,totaly1,'r','LineWidth',5)
% plot(totalxs,totalbinoprobs,'b')
% hold off

figure(22)
bar(totalxs,histftotaln)
xlim([-0.5,30])
title('histogram of marker counts in all full count FOVs, Poisson [red], Binomial [blue]')
hold on
plot(totalxs,totaly1,'r','LineWidth',5)
plot(totalxs,totalbinoprobs,'b')
hold off

figure(23)
bar(totalxs,histftotaln)
xlim([-0.5,30])
title('histogram of marker counts in all full count FOVs, Poisson [red]')
hold on
plot(totalxs,totaly1,'r','LineWidth',3)
hold off
% 
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%
% %Linear method calculations
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 
% Lconc=zeros(its,1);
% estimationerror=zeros(its,1);%To be used to check error of linear method
% estMxsqdiff=zeros(its,1);%To be used to calculate variance and then st dev of estimated Mx
% StockmarrEprop=zeros(its,1);%calculating Stockmarr's proportional error
% StockmarrEpropfpc=zeros(its,1);%calculating Stockmarr's proportional error, corrected for finite number
% lineffort=zeros(its,1);%Effort required for collecting linear data, Eqn 4.
% 
% meantxfpc=mean((Mx-ttotalx)/Mx)
% sqrt(meantxfpc)
% meantnfpc=mean((Mn-ttotaln)/Mn)
% sqrt(meantnfpc)
% 
% for i=1:its
%     Lestconci=ttotalx(i)*Mn/ttotaln(i);
%     StockmarrEprop(i)=sqrt( (1/ttotalx(i)) + (1/ttotaln(i)) );
%     txfpc=(Mx-ttotalx(i))/Mx;
%     tnfpc=(Mn-ttotaln(i))/Mn;
%     StockmarrEpropfpc(i)=sqrt( (1/ttotalx(i))*txfpc + (1/ttotaln(i))*tnfpc );
%     Lconc(i)=Lestconci;
%     Leffort(i)=(2*ttotalx(i)/calmeanx(i))+ttotalx(i)+ttotaln(i);
% end
% if its<=340
%     simsc4=sqrt(2/(its-1))*gamma(its/2)/gamma((its-1)/2);%Sampling st dev correction factor
% else
%     simsc4=1;
% end
% LstdTrueConc=sqrt(sum((Lconc-Mx).^2)/(its-1))/(Mx*simsc4);%Linear method st dev from true concentration
% LsimConcstd=std(Lconc)/(mean(Lconc)*simsc4);%Proportional standard deviation of simulated linear method concentration estimates
% 
% 
% 
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%
% %FOV method calculations
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%
% 
% FOVerror=zeros(its,1);%standard error of full count method (Eqn 3), using variance
% FOVeffort=zeros(its,1);%Effort required for collecting FOV data, Eqn 5.
% FOVerrorNorm=zeros(its,1);
% FOVeffort2=zeros(its,1);
% FOVerrorNorm2=zeros(its,1);
% 
% %finite population corrections
% meanxfpc=mean((Mx-caltotalx)/Mx)
% sqrt(meanxfpc)
% meannfpc=mean((Mn-ftotaln)/Mn)
% sqrt(meannfpc)
% 
% for i=1:its
%     fxfpc=(Mx-caltotalx(i))/Mx;
%     fnfpc=(Mn-ftotaln(i))/Mn;
%     FOVerrorNorm(i)=sqrt( (calsdx(i)/sqrt(fcal))^2 + (fsdn(i)/sqrt(fn))^2 );%Eqn 3
%     FOVerror(i)=sqrt( (calsdx(i)/sqrt(fcal))^2 + (sqrt(ftotaln(i))/ftotaln(i))^2 );%Eqn 3
%     FOVerrorNorm2(i)=sqrt( ((calsdx(i)/sqrt(fcal))^2)*fxfpc + ((fsdn(i)/sqrt(fn))^2)*fxfpc );%Eqn 3 with fpc
%     FOVerror2(i)=sqrt( ((calsdx(i)/sqrt(fcal))^2)*fxfpc + ((sqrt(ftotaln(i))/ftotaln(i))^2)*fxfpc );%Eqn 3 with fpc
%     FOVeffort(i)=2*fcal+caltotalx(i)+2*fn+ftotaln(i);
% end


% FstdTrueConc=sqrt(sum((Fconc-Mx).^2)/(its-1))/(Mx*simsc4);%FOV method st dev from true concentration
% FsimConcstd=std(Fconc)/(mean(Fconc)*simsc4);%Proportional standard deviation of simulated FOV method concentration estimates
% 100*mean(FOVerrorNorm)


% if fopt==1
%     fprintf(2,'XXXXXXXXXXXXX Data files XXXXXXXXXXXXXXXXXX\n')
% 
%     filenom1=strcat('.\SimData\LinConc_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
%     writematrix(Lconc,filenom1);
%     filenom2=strcat('.\SimData\LinError_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
%     writematrix(100*StockmarrEprop,filenom2);
%     filenom3=strcat('.\SimData\LinEffort_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
%     writematrix(Leffort,filenom3);
%     filenom4=strcat('.\SimData\FOVConc_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
%     writematrix(Fconc,filenom4);
%     filenom5=strcat('.\SimData\FOVError_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
%     writematrix(100*FOVerror,filenom5);
%     filenom6=strcat('.\SimData\FOVEffort_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
%     writematrix(FOVeffort,filenom6);
% end

% mean(ttotalx)
% (Mx-mean(ttotalx))/Mx
% (Mn-mean(ttotaln))/Mn
% mean(ttotaln)
% norecav=sqrt((1/mean(ttotalx))+(1/mean(ttotaln)))
% norecavFIN=sqrt((1/mean(ttotalx))*(Mx-mean(ttotalx))/Mx+(1/mean(ttotaln))*(Mn-mean(ttotaln))/Mn)
% 
% fprintf(2,'XXXXXXXXXXXXX Data for paper XXXXXXXXXXXXXXXXXX\n')
% fprintf('Lin: est. concentration = %f, st error (%%) = %f, sim std dev= %f\n',mean(Lconc),100*mean(StockmarrEprop),100*LsimConcstd)
% fprintf('Lin:, st error (%%) [no av over reciprocal] = %f, st error with fpc (%%) [no av over reciprocal] = %f, fpc st error (%%) = %f, \n',100*norecav,100*norecavFIN,100*mean(StockmarrEpropfpc))
% fprintf('Lin: prop. st dev from true concentration (%%) = %f, sampling effort = %f\n',100*LstdTrueConc,mean(Leffort))
% fprintf('FOV: est. concentration = %f, st error (%%) = %f, sim std dev= %f\n',mean(Fconc),100*mean(FOVerror),100*FsimConcstd)
% fprintf('FOV: st error (with sample std) = %f, fpc st error (%%) = %f, fpc st error (with sample std) = %f\n',100*mean(FOVerrorNorm),100*mean(FOVerror2),100*mean(FOVerrorNorm2))
% fprintf('FOV: prop. st dev from true concentration (%%) = %f, sampling effort = %f\n',100*FstdTrueConc,mean(FOVeffort))
% 
% fprintf('difference in mean conc. estimates = %f\n',mean(Lconc)-mean(Fconc))
% fprintf('XXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n')
% fprintf('two-sided, two-sample t-test that concentrations are different (assuming unequal variances):\n',mean(Lconc)-mean(Fconc))
% [result,p,ci,stats]=ttest2(Fconc,Lconc,'Vartype','unequal')

fprintf('XXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n')
% avLstE=100*mean(StockmarrEprop);
% avLsimE=100*LsimConcstd;
% avLtrueE=100*LstdTrueConc;
% avLeffort=mean(Leffort);
% avFstE=100*mean(FOVerror);
% avFsimE=100*FsimConcstd;
% avFtrueE=100*FstdTrueConc;
% avFeffort=mean(FOVeffort);
toc;
diary off
end
