function [N3star,fstar,FOVratio,FOVxdensity]=FOVoptimiserV1(Mx,Mn,omega,t,errormin,effortmax)
format compact;
format long;

aperturewidth=3;
apertureheight=3;
FOVarea=aperturewidth*apertureheight;

FOVxdensity=Mx*FOVarea/(100*100); %\bar{Y}_{3x}
rho=Mx/Mn;
FOVratio=rho*sqrt((omega+FOVxdensity)/(omega*rho+FOVxdensity));

sqrt1=sqrt(FOVxdensity+omega);
sqrt2=sqrt(FOVxdensity+rho*omega);
errorshift=(errormin/100)^2-t;

if (effortmax<0)&&(errormin>0)
    %fprintf(2,'Aiming for error = %f%%\n',errormin)
    N3star=(sqrt1+sqrt2)/(FOVxdensity*sqrt1*errorshift);
    fstar=N3star*FOVratio;
end

if (effortmax>0)&&(errormin<0)
    %fprintf(2,'Aiming for effort = %f units\n',effortmax)
    N3star=effortmax/(omega+FOVxdensity+sqrt1*sqrt2);
    fstar=N3star*FOVratio;
end

end