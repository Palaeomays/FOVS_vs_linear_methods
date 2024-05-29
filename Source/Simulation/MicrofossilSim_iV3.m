function [treturn,calreturn,fullreturn]=MicrofossilSim_iV3(Mx,Mn,tlim,fovsx,fovsn,printflag)

    dim=100;
    fossils=rand(Mx,2)*dim;
    exotics=rand(Mn,2)*dim;
    fossilsh=fossils(:,1);
    fossilsv=fossils(:,2);
    exoticsh=exotics(:,1);
    exoticsv=exotics(:,2);

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % Linear method
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    %Transect definition
    tstartpos=[5,83]; %coordinates of bottom left of transect
    tapheight=5;
    enough=0;
    while enough==0
        linearfossils=fossils(fossils(:,2) <= tstartpos(2)+tapheight & fossils(:,2) >= tstartpos(2),:);
        linearfossils=linearfossils(linearfossils(:,1)>=tstartpos(1),:);
        linearfossils=sortrows(linearfossils);
        nolfs=size(linearfossils,1); %total number of fossils that transect can see
        if nolfs<tlim
           tapheight=tapheight+1;
           tapheight+tstartpos(2);
        else
            enough=1;
        end
        if tapheight+tstartpos(2)>100
            fprintf(2,'nolfs=%d, tlim=%d\n',nolfs,tlim);
            error('XXXXXXXXX Not enough fossils in transect XXXXXXXXX');
        end
    end
    transfossils=linearfossils(1:min(tlim,nolfs),:);
    tnumberx=size(transfossils,1);
    tendpos=[transfossils(end,1),tstartpos(2)+tapheight];
    %fprintf('transect size: height= %f, width= %f\n',tapheight,tendpos(1)-tstartpos(1));
    
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
    Lpropn=Mn/tnumbern;
    
    %treturn of form [transect x count,std dev of x count,right edge of transect;transect n count, std dev of n count,number of counted n as proportion of total number on slide]
    treturn=[tnumberx,tsdx,tright;tnumbern,tsdn,Lpropn];
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % FOVs method --- Step 2, calibration counts
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %Fields of view (fovs) for calibration counts (x)
    fx=size(fovsx,2);%Number of calibration FOVs, N_3 in Eqn 3.
    fxcounts=zeros(1,fx);
    aperturewidthx=fovsx(2,1)-fovsx(1,1);
    apertureheightx=fovsx(4,1)-fovsx(3,1);
    fovxarea=aperturewidthx*apertureheightx;
    %Count number of fossils in each fov
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
    fmeanx=mean(fxcounts);%Mean number of fossils per FOV in calibration
    fvarx=var(fxcounts)/(fmeanx^2);%Proportional sample variance of mean number of fossils per FOV in calibration
    calc4=sqrt(2/(fx-1))*gamma(fx/2)/gamma((fx-1)/2);%Sampling st dev correction factor for calibration counts
    fovsdx=(std(fxcounts)/fmeanx)/calc4;%Proportional sample standard deviation, corrected with c4, SP3 in Eqn 3.
    
    %calreturn of form [counted x in all calibration FOVs,mean x in calibration FOVs, std dev of counted x in calibration FOVs]
    calreturn=[fnumberx,fmeanx,fvarx,fovsdx];

    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % FOVs method --- Step 4, full counts
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    %Fields of view (fovs) for full counts (n)
    fn=size(fovsn,2);
    fncounts=zeros(1,fn);
    truefxcounts=zeros(1,fn);
    aperturewidthn=fovsn(2,1)-fovsn(1,1);
    apertureheightn=fovsn(4,1)-fovsn(3,1);
    fovnarea=aperturewidthn*apertureheightn;
    %Count number of exotics and fossils in each fov
    for i= 1:fn
        %find edges of fov
        fovleft=fovsn(1,i);
        fovright=fovsn(2,i);
        fovbot=fovsn(3,i);
        fovtop=fovsn(4,i);
        
        %fossils
        fossilhcheck=and(fossils(:,1)>=fovleft,fossils(:,1)<=fovright);
        fossilvcheck=and(fossils(:,2)>=fovbot,fossils(:,2)<=fovtop);
        fossilcheck=and(fossilhcheck,fossilvcheck);
        fcountedx=[fossilsh(fossilcheck),fossilsv(fossilcheck)];
        truefxcounts(i)=size(fcountedx,1);
        
        %exotics
        exotichcheck=and(exotics(:,1)>=fovleft,exotics(:,1)<=fovright);
        exoticvcheck=and(exotics(:,2)>=fovbot,exotics(:,2)<=fovtop);
        exoticcheck=and(exotichcheck,exoticvcheck);
        fcountedn=[exoticsh(exoticcheck),exoticsv(exoticcheck)];
        fncounts(i)=size(fcountedn,1);
    end
    
    fovc4=sqrt(2/(fn-1))*gamma(fn/2)/gamma((fn-1)/2);%Sampling st dev correction factor for full counts
    
    fnumbertruex=sum(truefxcounts); %Actual number of total x in full count FOVs
    fmeantruex=mean(truefxcounts); 
    fovsdtruex=(std(truefxcounts)/fmeantruex)/fovc4;
    fnumbern=sum(fncounts);
    fmeann=mean(fncounts);
    fovsdn=(std(fncounts)/fmeann)/fovc4;
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % FOVs method --- Step 5, concentration and error
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    xhat=fmeanx*fn; %Estimated number of total x in full count FOVs
    concentration= xhat*Mn/fnumbern; %Estimated total number of x in slade (FOV method)
    
    %fullreturn of form:
    %row 1= [counted x in all full FOVs, mean x in full FOVs, std dev of x in full FOVs, estimated total x]
    %row 2= [counted n in all full FOVs, mean n in full FOVs, std dev of n in full FOVs, concentration]
    fullreturn=[fnumbertruex,fmeantruex,fovsdtruex,xhat;fnumbern,fmeann,fovsdn,concentration];
    
    %print the fossils and exotics with the transect and the fovs
    if printflag==1
        hold off
        figure(1)
        scatter(fossils(:,1),fossils(:,2),20,'.')
        xlim([0,dim])
        ylim([0,dim])
        pbaspect([1 1 1])
        hold on
        scatter(exotics(:,1),exotics(:,2),12,'d','filled','r')
        rectangle('Position',[tstartpos(1),tstartpos(2),twidth,theight],'EdgeColor',[0.2660 0.6740 0.2880],'LineWidth',2,'FaceColor',[0.2660 0.6740 0.2880])
        for i=1:fx
            rectangle('Position',[fovsx(1,i),fovsx(3,i),aperturewidthx,apertureheightx],'FaceColor','b');
        end
        for i=1:fn
            rectangle('Position',[fovsn(1,i),fovsn(3,i),aperturewidthn,apertureheightn],'EdgeColor','m','FaceColor','m');
        end
        hold off
                      
        %Zoomed in figure on one calibration FOV
        figure(2)
        scatter(fossils(:,1),fossils(:,2),300,'.')
        xlim([fovsx(1,1)*0.99,fovsx(2,1)*1.01])
        ylim([fovsx(3,1)*0.99,fovsx(4,1)*1.01])
        pbaspect([1 1 1])
        hold on
        scatter(exotics(:,1),exotics(:,2),100,'filled','d')
        rectangle('Position',[fovsx(1,1),fovsx(3,1),aperturewidthx,apertureheightx]);
    end
    hold off
end
