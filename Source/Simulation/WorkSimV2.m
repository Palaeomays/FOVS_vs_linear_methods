function [LPrecWorkOpt,FPrecWorkOpt] = WorkSimV2(Mx,Mn,tmax,fx,fn,its,fileopt)
fprintf(2,'Mx= %d, Mn= %d, fx= %d, fn= %d, its= %d\nStarted at %s\n',Mx,Mn,fx,fn,its,datestr(clock))
noprints=10;%The number of print statements to produce while the code is running to record the time taken
format compact;
format long;

% building the calibration fields of view
% fcal=21; %Number of calibration fields of view
% aperturewidth=3;
% apertureheight=3;
% FOVarea=aperturewidth*apertureheight;
% leftsx=[3.5:9:99,48.5*ones(1,10)];
% bottomsx=[48.5*ones(1,11),4.5,13.5,22.5,31.5,40.5,56.5,65.5,74.5,83.5,92.5];
% 
% take only fx of the calibration FOVs
% leftsx=leftsx(1:fx);
% bottomsx=bottomsx(1:fx);
% 
% rightsx=leftsx+aperturewidth;
% topsx=bottomsx+apertureheight;
% fovsx=[leftsx;rightsx;bottomsx;topsx];%fields of view with format [left edge; right edge; bottom edge; top edge]
% 
% figure(100)
% cla;
% for i=1:fcal
%     rectangle('Position',[fovsx(1,i),fovsx(3,i),aperturewidth,apertureheight],'FaceColor','b');
% end
% xlim([0,100])
% ylim([0,100])
% pbaspect([1 1 1])
% 
% building the full count fields of view
% block=floor(fn/50);%First 50 FOVs in left block, then next 50 FOVs in right block
% nonews=mod(fn,50);%Number of FOVs in the new block
% newlefts=zeros(1,nonews);%List of left edges of FOVs in new block
% newbottoms=zeros(1,nonews);%List of bottom edges of FOVs in new block
% 
% for fj=1:nonews
%      row=floor((fj-1)/5);
%      if fj<=25
%          botj=2+10*row;
%      else
%          botj=5+10*row;
%      end
%      leftj=2+10*mod(fj-1,5);
%      if block==1
%          leftj=leftj+53;
%      end
%      newlefts(fj)=leftj;
%      newbottoms(fj)=botj;
% end
% 
% if block==1
%     leftsn=[repmat(2:10:42,1,10),newlefts];
%     bottomsn=[2*ones(1,5),12*ones(1,5),22*ones(1,5),32*ones(1,5),42*ones(1,5),55*ones(1,5),65*ones(1,5),75*ones(1,5),85*ones(1,5),95*ones(1,5)];
%     bottomsn=[bottomsn,newbottoms];
% else
%     leftsn=newlefts;
%     bottomsn=newbottoms;
% end
% 
% rightsn=leftsn+aperturewidth;
% topsn=bottomsn+apertureheight;
% fovsn=[leftsn;rightsn;bottomsn;topsn];%fields of view with format [left edge; right edge; bottom edge; top edge]

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

length(allFOVs)
length(calFOVs)
length(fullFOVs)

tic;

for i=1:its
    if i==1
        [WorkPrecFAll,WorkPrecLAll]=WorkSimV2_i(Mx,Mn,tmax,calFOVs,fullFOVs);
    else
        [WorkPrecF,WorkPrecL]=WorkSimV2_i(Mx,Mn,tmax,calFOVs,fullFOVs);
        
        WorkPrecLAll(2,:)=WorkPrecLAll(2,:)+WorkPrecL(2,:);%Number of exotics in transect (Linear)
        WorkPrecLAll(3,:)=WorkPrecLAll(3,:)+WorkPrecL(3,:);%Effort (Linear)
        WorkPrecLAll(4,:)=WorkPrecLAll(4,:)+WorkPrecL(4,:);%Precision (Linear)
        
        WorkPrecFAll(3,:)=WorkPrecFAll(3,:)+WorkPrecF(3,:);%Precision (FOVs)
        WorkPrecFAll(4,:)=WorkPrecFAll(4,:)+WorkPrecF(4,:);%Precision (FOVs)
        WorkPrecFAll(5,:)=WorkPrecFAll(5,:)+WorkPrecF(5,:);%Calibration counts (FOVs)
        WorkPrecFAll(6,:)=WorkPrecFAll(6,:)+WorkPrecF(6,:);%Full counts (FOVs)
    end
 
    if mod(i,floor(its/noprints))==0
        fprintf('i= %d, at %s\n',i,datestr(clock))
    end
end
WorkPrecLAll(2,:)=WorkPrecLAll(2,:)/its;
WorkPrecLAll(3,:)=WorkPrecLAll(3,:)/its;
WorkPrecLAll(4,:)=WorkPrecLAll(4,:)/its;

WorkPrecFAll(3,:)=WorkPrecFAll(3,:)/its;
WorkPrecFAll(4,:)=WorkPrecFAll(4,:)/its;
WorkPrecFAll(5,:)=WorkPrecFAll(5,:)/its;
WorkPrecFAll(6,:)=WorkPrecFAll(6,:)/its;

WorkPrecFAll(isinf(WorkPrecFAll))=NaN;%All indeterminate values set to NaN, so that they can be excluded from a mean calculation
WorkPrecLAll(isinf(WorkPrecLAll))=NaN;%All indeterminate values set to NaN, so that they can be excluded from a mean calculation


epsilon=2;%Accuracy cutoff for precision

%WorkPrecL of form: [no. fossils in transect; no. markers in transect; effort; precision]
ConcL=zeros(3,size(WorkPrecLAll,2));

ConcL(1,:)=WorkPrecLAll(3,:);
ConcL(2,:)=Mn*WorkPrecLAll(1,:)./WorkPrecLAll(2,:);
ConcL(3,:)=100*(ConcL(2,:)-Mx)/Mx;

figure(100)
scatter(ConcL(1,:),ConcL(3,:))
ylim([-1.5*epsilon,1.5*epsilon])

%Linear method checks
transratio=mean(WorkPrecLAll(1,:)./WorkPrecLAll(2,:))
actualratio=Mx/Mn

%fprintf(2,'final Linear Conc est=%f, %% error=%f\n',ConcL(2,end),ConcL(3,end));

%WorkPrecF of form: [no. calibration FOVs; no. full count FOVs; effort; precision;no. fossils in cal count (fixed);no. exotics in full count]
%FOVS concentration estimates
ConcF=zeros(3,size(WorkPrecFAll,2));

Y3bar=WorkPrecFAll(5,end)/WorkPrecFAll(1,end);
xhats=Y3bar*WorkPrecFAll(2,:);
propns=Mn./WorkPrecFAll(6,:);
ConcF(1,:)=WorkPrecFAll(3,:);%FOVS effort
ConcF(2,:)=xhats.*propns;%FOVS concentration estimates
ConcF(3,:)=100*(ConcF(2,:)-Mx)/Mx;%FOVS %error in concentration estimate

figure(101)
scatter(ConcF(1,:),ConcF(3,:))
ylim([-1.5*epsilon,1.5*epsilon])

%FOVS method checks
calprob=fx*FOVarea/(100*100);
expcalcount=calprob*Mx
actualcalcount=WorkPrecFAll(5,end)

fullprob=fn*FOVarea/(100*100);
expfullcount=fullprob*Mn
actualfullcount=WorkPrecFAll(6,end)

%fprintf(2,'final FOVS Conc est=%f, %% error=%f\n',ConcF(2,end),ConcF(3,end));

%Work graph
ConcLstartindex=-1;
ConcLstartwork=-1;

for i=1:size(ConcL,2)
    noleft=size(ConcL,2)-i+1;
    deltas=sum(abs(ConcL(3,i:end))<epsilon);
    if deltas==noleft
        ConcLstartindex=i;
        ConcLstartwork=ConcL(1,i);
        break;
    end
end
% ConcLstartindex
% ConcLstartwork

ConcFstartindex=-1;
ConcFstartwork=-1;

for i=1:size(ConcF,2)
    noleft=size(ConcF,2)-i+1;
    deltas=sum(abs(ConcF(3,i:end))<epsilon);
    if deltas==noleft
        ConcFstartindex=i;
        ConcFstartwork=ConcF(1,i);
        break;
    end
end
% ConcFstartindex
% ConcFstartwork

% figure(200)
% scatter(WorkPrecFAll(3,:),WorkPrecFAll(4,:))
% xline(ConcFstartwork,'-b');
% xlim([0,1.1*max(WorkPrecFAll(3,:))])
% ylim([0,1.1*max(WorkPrecFAll(4,:))])
% hold on
% scatter(WorkPrecLAll(3,:),WorkPrecLAll(4,:))
% xline(ConcLstartwork,'-r');
% hold off

%LPrecWorkOpt of form: [no. fossils in transect; no. markers in transect; effort; precision; concentration estimate; % error in concentration estimate; effort beyond which accuracy goal is met]
Lstart=ConcLstartwork*ones(1,size(ConcL,2));
LPrecWorkOpt=[WorkPrecLAll;ConcL(2:3,:);Lstart];

%FPrecWorkOpt of form: [no. calibration FOVs; no. full count FOVs; effort; precision;no. fossils in cal count (fixed);no. exotics in full count; concentration estimate; % error in concentration estimate; effort beyond which accuracy goal is met]
Fstart=ConcFstartwork*ones(1,size(ConcF,2));
FPrecWorkOpt=[WorkPrecFAll;ConcF(2:3,:);Fstart];

if fileopt==1
    filenom1=strcat('.\SimData\LPrecWork_Mx',int2str(Mx),'_Mn',int2str(Mn),'_tmax',int2str(tmax),'_its',int2str(its),'.csv')
    writematrix(LPrecWorkOpt,filenom1);
    filenom2=strcat('.\SimData\FPrecWork_Mx',int2str(Mx),'_Mn',int2str(Mn),'_fx',int2str(fx),'_fn',int2str(fn),'_its',int2str(its),'.csv')
    writematrix(FPrecWorkOpt,filenom2);
end

toc;
end
