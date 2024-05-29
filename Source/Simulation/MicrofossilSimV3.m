function [avLstE,avLsimE,avLtrueE,avLeffort,avFstE,avFsimE,avFtrueE,avFeffort,pError] = MicrofossilSimV3(Mx,Mn,t,tlim,fx,fn,its,fopt)
noprints=10;%The number of print statements to produce while the code is running to record the time taken
format compact;
format long;

%building all fields of view
aperturewidth=3;
apertureheight=3;
FOVarea=aperturewidth*apertureheight;

FOVlefts=[3.5:5:100-6.5];
noHFOVs=length(FOVlefts);
FOVlefts=repmat(FOVlefts,1,length(FOVlefts));
FOVrights=FOVlefts+3;

%fields of view with format [left edge; right edge; bottom edge; top edge]
allFOVs=[FOVlefts;FOVrights;zeros(1,length(FOVlefts));zeros(1,length(FOVlefts))];

if fx+fn>length(allFOVs)
    fprintf(2,'fx+fn = %d, total FOVS = %d\n',fx+fn,length(allFOVs))
    error('XXX Not enough total fields of view XXX');
end

for i=1:noHFOVs
    startpos=(i-1)*noHFOVs+1;
    endpos=(i)*noHFOVs;
    allFOVs(3,startpos:endpos)=FOVlefts(i);
    allFOVs(4,startpos:endpos)=FOVlefts(i)+apertureheight;
end

calFOVs=allFOVs(:,1:fx);
fullFOVs=allFOVs(:,(fx+1):fn+fx);


ttotalx=zeros(its,1); %total x in transect (linear)
ttotaln=zeros(its,1); %total n in transect (linear)
tsdx=zeros(its,1); %st dev of x in transect, sqrt(x)/meanx(x) (linear)
tsdn=zeros(its,1); %st dev of n in transect, sqrt(n)/meanx(n) (linear)
transrights=zeros(its,1); %Number of steps needed to find transect. Only used to evaluate efficiency of program.
tpropns=zeros(its,1); %Proportion of marker grains on slide counted

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
%treturn of form [transect x count,std dev of x count,right edge of transect;transect n count, std dev of n count,number of counted n as proportion of total number on slide]
%calreturn of form [counted x in all calibration FOVs,mean x in calibration FOVs, std dev of counted x in calibration FOVs]
%fullreturn of form:
    %row 1= [counted x in all full FOVs, mean x in full FOVs, std dev of x in full FOVs, estimated total x]
    %row 2= [counted n in all full FOVs, mean n in full FOVs, std dev of n in full FOVs, concentration]
for i=1:its
    if i==its
        [treturn,calreturn,fullreturn]=MicrofossilSim_iV3(Mx,Mn,tlim,calFOVs,fullFOVs,1);
    else
        [treturn,calreturn,fullreturn]=MicrofossilSim_iV3(Mx,Mn,tlim,calFOVs,fullFOVs,0);
    end
    
    transrights(i)=treturn(1,3); %Right edge of transect
    ttotalx(i)=treturn(1,1);%total x in transect (linear)
    ttotaln(i)=treturn(2,1);%total n in transect (linear)
    tsdx(i)=treturn(1,2);%st dev of x in transect, sqrt(x)/meanx(x) (linear)
    tsdn(i)=treturn(2,2);%st dev of n in transect, sqrt(n)/meanx(n) (linear)
    tpropns(i)=treturn(2,3);%proportion of marker grains counted in transect (linear)
    
    caltotalx(i)=calreturn(1); %calibration count of x in all FOVs
    calmeanx(i)=calreturn(2); %calibration mean number of x per FOV
    calvarx(i)=calreturn(3);
    calsdx(i)=calreturn(4); %calibration st dev of mean number of x (proportional, corrected with c4)

    ftotalx(i)=fullreturn(1,1); %full count of x in all FOVs (explicit counting, as opposed to estimating using mean from calibration)
    fmeanx(i)=fullreturn(1,2); %mean number of x in full count (explicit mean, as opposed to using mean from calibration)
    fsdx(i)=fullreturn(1,3); %st dev of mean number of x in full count (explicit st dev, as opposed to using st dev from calibration
    ftotaln(i)=fullreturn(2,1); %count of exotics in full count in all FOVs
    fmeann(i)=fullreturn(2,2); %mean number of exotics per FOV
    fsdn(i)=fullreturn(2,3); %st dev of mean number of exotics per FOV
    
    xhats(i)=fullreturn(1,4); %Estimated number of total fossils in full count (using calibration count)
    Fconc(i)=fullreturn(2,4); %Concentration calculation (assuming V=1)
    %fprintf('time for i = %d is %f secs at %s\n',i,toc,datestr(clock));
    if mod(i,floor(its/noprints))==0
        fprintf('i= %d, at %s\n',i,datestr(clock))
    end
end

talln=sum(ttotaln)
falln=sum(ftotaln)

%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Linear method calculations
%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Lconc=zeros(its,1);
estimationerror=zeros(its,1);%To be used to check error of linear method
estMxsqdiff=zeros(its,1);%To be used to calculate variance and then st dev of estimated Mx
StockmarrE=zeros(its,1);%calculating Stockmarr's proportional error, multiplied by counted x
StockmarrEprop=zeros(its,1);%calculating Stockmarr's proportional error
StockmarrEpropfpc=zeros(its,1);%calculating Stockmarr's proportional error, corrected for finite number
lineffort=zeros(its,1);%Effort required for collecting linear data, Eqn 4.
for i=1:its
    Lestconci=ttotalx(i)*Mn/ttotaln(i);
    StockmarrEprop(i)=sqrt( t+ (1/ttotalx(i)) + (1/ttotaln(i)) );
    Lconc(i)=Lestconci;
    Leffort(i)=(2*ttotalx(i)/calmeanx(i))+ttotalx(i)+ttotaln(i);
    %Finite size simulation correction
    txfpc=(Mx-ttotalx(i))/Mx;
    tnfpc=(Mn-ttotaln(i))/Mn;
    StockmarrEpropfpc(i)=sqrt(t+ (1/ttotalx(i))*txfpc + (1/ttotaln(i))*tnfpc );
end
if its<=340
    simsc4=sqrt(2/(its-1))*gamma(its/2)/gamma((its-1)/2);%Sampling st dev correction factor
else
    simsc4=1;
end
LstdTrueConc=sqrt(t+ sum((Lconc-Mx).^2)/((its-1)*Mx^2*simsc4^2));%Linear method st dev from true concentration
meanLprop=Mn*its/sum(ttotaln);%mean reciprocal proportion of counted n (Linear method), using unbiased estimator of the reciprocal
avLconc=mean(ttotalx)*meanLprop;%mean concentration estimate (Linear method)
LsimConcstd=std(Lconc)/(avLconc*simsc4);%Proportional standard deviation of simulated linear method concentration estimates

%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%FOV method calculations
%%%%%%%%%%%%%%%%%%%%%%%%%%%%

FOVerror=zeros(its,1);%standard error of full count method (Eqn 3)
FOVerror2=zeros(its,1);%standard error of full count method (Eqn 3), but with sample std dev instead of sqrt(n)/n
FOVerrorfpc=zeros(its,1);%standard error of full count method (Eqn 3)
FOVeffort=zeros(its,1);%Effort required for collecting FOV data, Eqn 5.

for i=1:its
    FOVerror(i)=sqrt(t+ (calsdx(i)/sqrt(fx))^2 + (sqrt(ftotaln(i))/ftotaln(i))^2 );%Eqn 3
    FOVerror2(i)=sqrt(t+ (calsdx(i)/sqrt(fx))^2 +(fsdn(i)/sqrt(fn))^2 );%Eqn 3, but with sample std dev instead of sqrt(n)/n
    FOVeffort(i)=(2*fx)+caltotalx(i)+(2*fn)+ftotaln(i);
    %Finite size simulation correction
    fxfpc=(Mx-caltotalx(i))/Mx;
    fnfpc=(Mn-ftotaln(i))/Mn;
    FOVerrorfpc(i)=sqrt(t+ ((calsdx(i)/sqrt(fx))^2)*fxfpc + ((sqrt(ftotaln(i))/ftotaln(i))^2)*fnfpc);%Eqn 3 with finite population correction
end
FstdTrueConc=sqrt(t+ sum((Fconc-Mx).^2)/((its-1)*Mx^2*simsc4^2));%FOV method st dev from true concentration

meanFprop=Mn*its/sum(ftotaln);%mean reciprocal proportion of counted n (FOV method), using unbiased estimator of the reciprocal
avFconc=mean(xhats)*meanFprop;%mean concentration estimate (FOV method)

FsimConcstd=std(Fconc)/(avFconc*simsc4);%Proportional standard deviation of simulated FOV method concentration estimates

if fopt==1
    fprintf(2,'XXXXXXXXXXXXX Data files XXXXXXXXXXXXXXXXXX\n')

    filenom1=strcat('.\SimData\LinConc_Mx',int2str(Mx),'_Mn',int2str(Mn),'_T',int2str(10^4*t),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(Lconc,filenom1);
    filenom2=strcat('.\SimData\LinError_Mx',int2str(Mx),'_Mn',int2str(Mn),'_T',int2str(10^4*t),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(100*StockmarrEprop,filenom2);
    filenom3=strcat('.\SimData\LinEffort_Mx',int2str(Mx),'_Mn',int2str(Mn),'_T',int2str(10^4*t),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(Leffort,filenom3);
    filenom4=strcat('.\SimData\FOVConc_Mx',int2str(Mx),'_Mn',int2str(Mn),'_T',int2str(10^4*t),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(Fconc,filenom4);
    filenom5=strcat('.\SimData\FOVError_Mx',int2str(Mx),'_Mn',int2str(Mn),'_T',int2str(10^4*t),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(100*FOVerror,filenom5);
    filenom6=strcat('.\SimData\FOVEffort_Mx',int2str(Mx),'_Mn',int2str(Mn),'_T',int2str(10^4*t),'_tlim',int2str(tlim),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(FOVeffort,filenom6);
end

aveffort=(mean(Leffort)+mean(FOVeffort))/2;
Lworkscale=mean(Leffort)/aveffort
Fworkscale=mean(FOVeffort)/aveffort

fprintf(2,'XXXXXXXXXXXXX Data for paper V2 XXXXXXXXXXXXXXXXXX\n')

fprintf('Lin: est. concentration = %f, sampling effort = %f, scaled (\\tilde{\\sigma}_L, Eqn7) st error (%%) = %f\n',avLconc, mean(Leffort),100*mean(StockmarrEprop)*Lworkscale)
fprintf('Lin: scaled fpc (\\hat{\\sigma}_L, Eqn11) st error (%%) = %f, prop. st dev from true concentration (%%) = %f\n',100*mean(StockmarrEpropfpc)*Lworkscale,100*LstdTrueConc*Lworkscale)
fprintf('Lin: %% diff between True error and st error (fpc) = %f\n',100*(LstdTrueConc-mean(StockmarrEpropfpc))/LstdTrueConc)
fprintf('FOV: est. concentration = %f, sampling effort = %f, scaled (\\tilde{\\sigma}_F, Eqn8) st error (%%) = %f\n',avFconc, mean(FOVeffort),100*mean(FOVerror)*Fworkscale)
fprintf('FOV: scaled fpc (\\hat{\\sigma}_F, Eqn12) st error (%%) = %f, prop. st dev from true concentration (%%) = %f\n',100*mean(FOVerrorfpc)*Fworkscale,100*FstdTrueConc*Fworkscale)
fprintf('FOV: %% diff between True error and st error (fpc) = %f\n\n',100*(FstdTrueConc-mean(FOVerrorfpc))/FstdTrueConc)


fprintf('XXXXXXXXXXXXX Simulation stats V2 XXXXXXXXXXXXXXXXXX\n')
fprintf('average work for both methods = %f\n',aveffort)
fprintf('difference in conc. estimates = %f, %% difference in conc. estimates = %f\n',avLconc-avFconc,100*(avLconc-avFconc)/avFconc)
fprintf('Lin: %% error of estimate from true conc. = %f\n',(avLconc-Mx)/Mx)
fprintf('Lin: %% error between infinite and finite (std error) = %f, sim std dev= %f\n',100*(mean(StockmarrEprop)-mean(StockmarrEpropfpc))/mean(StockmarrEpropfpc),100*LsimConcstd)
fprintf('Lin: %% diff between True error and st error (no fpc) = %f,\n',100*(LstdTrueConc-mean(StockmarrEprop))/LstdTrueConc)
fprintf('FOV: %% error of estimate from true conc. = %f\n',(avFconc-Mx)/Mx)
fprintf('FOV: %% error between infinite and finite (std error) = %f, st error (%%) using std(n) = %f, sim std dev= %f\n',100*(mean(FOVerror)-mean(FOVerrorfpc))/mean(FOVerrorfpc),100*mean(FOVerror2),100*FsimConcstd)
fprintf('FOV: %% diff between True error and st error (no fpc)= %f\n\n',100*(FstdTrueConc-mean(FOVerror))/FstdTrueConc)


fprintf('paired t-test on concentrations:\n')
[resultConc,pConc,ciConc,statsConc]=ttest(Fconc,Lconc)

fprintf('Wilcoxon signed rank test on errors:\n')
[pError,resultError,statsError]=signrank(100*StockmarrEprop,100*FOVerror)


avLstE=100*mean(StockmarrEprop);
avLsimE=100*LsimConcstd;
avLtrueE=100*LstdTrueConc;
avLeffort=mean(Leffort);
avFstE=100*mean(FOVerror);
avFsimE=100*FsimConcstd;
avFtrueE=100*FstdTrueConc;
avFeffort=mean(FOVeffort);
toc;
fprintf('XXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n\n\n')
end
