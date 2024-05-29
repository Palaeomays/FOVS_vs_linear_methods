function [WorkPrecF,WorkPrecL]=WorkSimV2_i(Mx,Mn,tmax,fovsx,fovsn)

    dim=100;
    fossils=rand(Mx,2)*dim;
    exotics=rand(Mn,2)*dim;
    fossilsh=fossils(:,1);
    fossilsv=fossils(:,2);
    exoticsh=exotics(:,1);
    exoticsv=exotics(:,2);
    tab=0;

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % FOVs method --- Step 2, calibration counts
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %Fields of view (fovs) for calibration counts (x)
    fx=size(fovsx,2);%Number of calibration FOVs, N_3 in Eqn 3.
    fxcounts=zeros(1,fx);
    aperturewidthx=fovsx(2,1)-fovsx(1,1);
    apertureheightx=fovsx(4,1)-fovsx(3,1);
    fovxarea=aperturewidthx*apertureheightx;
    %Count number of fossils in each calibration fov
    for i= 1:fx
        fovleft=fovsx(1,i);
        fovright=fovsx(2,i);
        fossilhcheck=and(fossils(:,1)>=fovleft,fossils(:,1)<=fovright);
        fovbot=fovsx(3,i);
        fovtop=fovsx(4,i);
        fossilvcheck=and(fossils(:,2)>=fovbot,fossils(:,2)<=fovtop);
        fossilcheck=and(fossilhcheck,fossilvcheck);
        fcountedx=[fossilsh(fossilcheck),fossilsv(fossilcheck)];
        fxcounts(i)=size(fcountedx,1); 
    end
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % FOVs method --- Step 3, statistics of calibration counts
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    fnumberx=sum(fxcounts);%Total number of fossils counted in calibration
    fmeanx=mean(fxcounts);%Mean number of fossils per FOV in calibration (called \overline{Y}_3)
    calc4=sqrt(2/(fx-1))*gamma(fx/2)/gamma((fx-1)/2);%Sampling st dev correction factor for calibration counts
    fovsdx=(std(fxcounts)/fmeanx)/calc4;%Proportional sample standard deviation, corrected with c4, SP3 in Eqn 3.
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % FOVs method --- Step 4, full counts
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %Fields of view (fovs) for full counts (n)
    fn=size(fovsn,2);
    fncounts=zeros(1,fn);
    
    %WorkPrecF of form: [no. calibration FOVs; no. full count FOVs; effort; precision;no. fossils in cal count (fixed);no. exotics in full count]
    WorkPrecF=zeros(6,fn);
    WorkPrecF(1,:)=fx;
    
    aperturewidthn=fovsn(2,1)-fovsn(1,1);
    apertureheightn=fovsn(4,1)-fovsn(3,1);
    fovnarea=aperturewidthn*apertureheightn;
    %Count number of exotics and fossils in each fov
    ncounter=0;
    for i= 1:fn
        %find edges of fov
        fovleft=fovsn(1,i);
        fovright=fovsn(2,i);
        fovbot=fovsn(3,i);
        fovtop=fovsn(4,i);
        
        %exotics
        exotichcheck=and(exotics(:,1)>=fovleft,exotics(:,1)<=fovright);
        exoticvcheck=and(exotics(:,2)>=fovbot,exotics(:,2)<=fovtop);
        exoticcheck=and(exotichcheck,exoticvcheck);
        fcountedn=[exoticsh(exoticcheck),exoticsv(exoticcheck)];
        fnumberni=size(fcountedn,1);
        fncounts(i)=fnumberni;
        ncounter=ncounter+fnumberni;
        WorkPrecF(2,i)=i;
        WorkPrecF(3,i)=2*fx+fnumberx+2*i+ncounter; %eF
        WorkPrecF(4,i)=100*sqrt( tab+ (fovsdx/sqrt(fx))^2 + (1/ncounter) );%Eqn 3
        WorkPrecF(5,i)=fnumberx;
        WorkPrecF(6,i)=ncounter;
    end
            
 
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % Linear method
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %Transect definition
    tstartpos=[5,60]; %coordinates of bottom left of transect
    tapheight=5;
    linearfossils=fossils(fossils(:,2) <= tstartpos(2)+tapheight & fossils(:,2) >= tstartpos(2),:);
    linearfossils=linearfossils(linearfossils(:,1)>=tstartpos(1),:);
    linearfossils=sortrows(linearfossils);
    nolfs=size(linearfossils,1); %total number of fossils that transect can see
    if nolfs<tmax
        fprintf(2,'nolfs=%d, tmax=%d\n',nolfs,tmax);
        error('XXXXXXXXX Not enough fossils in transect XXXXXXXXX');
    end
    tmin=50;%Number of fossils to count in smallest transect
    tstep=10;%Number of fossils to increase size of transect
    ts=tmin:tstep:tmax;
    nots=length(ts);
    
    %WorkPrecL of form: [no. fossils in transect; no. markers in transect; effort; precision]
    WorkPrecL=zeros(4,nots);
        
    for i=1:nots
        tlim=ts(i);
       
        transfossils=linearfossils(1:min(tlim,nolfs),:);
        tnumberx=size(transfossils,1);
        tendpos=[transfossils(end,1),tstartpos(2)+tapheight];

        %fossil statistics
        tsdx=sqrt(tnumberx)/tnumberx;
        twidth=tendpos(1)-tstartpos(1);
        theight=tendpos(2)-tstartpos(2);
        transectarea=twidth*theight;    

        %Counting exotics in transect
        tleft=tstartpos(1);
        tbot=tstartpos(2);
        tright=tendpos(1);
        ttop=tendpos(2);
        exotichcheck=and(tleft<=exotics(:,1),exotics(:,1)<=tright);%binary vector marking exotics in horizontal dimension of transect
        exoticvcheck=and(tbot<=exotics(:,2),exotics(:,2)<=ttop);%binary vector marking exotics in vertical dimension of transect
        exoticcheck=and(exotichcheck,exoticvcheck);%binary vector marking exotics in transect
        exoticsh=exotics(:,1);
        exoticsv=exotics(:,2);
        tcountedn=sortrows([exoticsh(exoticcheck),exoticsv(exoticcheck)]);%List of exotics found in transect
        tnumbern=size(tcountedn,1);
        tsdn=sqrt(tnumbern)/tnumbern;
        WorkPrecL(1,i)=tlim;
        WorkPrecL(2,i)=tnumbern;
        WorkPrecL(3,i)=(2*tlim/fmeanx)+tlim+tnumbern;
        WorkPrecL(4,i)=100*sqrt(tab+ (1/tlim)+(1/tnumbern));
    end
    
end
