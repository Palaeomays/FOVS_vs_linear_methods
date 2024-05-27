function [N3star,fstar,FOVratio,FOVxdensity]=FOVoptimiserV1(Mx,Mn,alpha,t,errormin,workmax)
format compact;
format long;

aperturewidth=3;
apertureheight=3;
FOVarea=aperturewidth*apertureheight;

FOVxdensity=Mx*FOVarea/(100*100); %Y3bar
rho=Mx/Mn;
FOVratio=rho*sqrt((alpha+FOVxdensity)/(alpha*rho+FOVxdensity));

sqrt1=sqrt(FOVxdensity+alpha);
sqrt2=sqrt(FOVxdensity+rho*alpha);
errorshift=(errormin/100)^2-t;

if (workmax<0)&&(errormin>0)
    %fprintf(2,'Aiming for error = %f%%\n',errormin)
    N3star=(sqrt1+sqrt2)/(FOVxdensity*sqrt1*errorshift);
    fstar=N3star*FOVratio;
end

if (workmax>0)&&(errormin<0)
    %fprintf(2,'Aiming for work = %f units\n',workmax)
    N3star=workmax/(alpha+FOVxdensity+sqrt1*sqrt2);
    fstar=N3star*FOVratio;
end

end