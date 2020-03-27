%% Sample from the original xlsx file and generate several new xlsx files
%
% RewriteXLSX(originalfile, sheetname, Nnewfile, target, Ntarget, savedir, newfilename)
%
% - originalfile
%   string, the name of the file to be sampled
% - sheetname
%   string, the name of the sheet to be sampled
% - Nnewfile
%   integer, denotes the number of new files
% - target
%   string, 'row'/'col', denotes the target of sampling
% - Ntarget
%   integer, denotes the number of rows/columns per file
% - savedir
%   string, save directory
% - newfilename
%   string, the cap of the names of the new files
%
% Example:
% RewriteXLSX('hahaha','A',30,'row',20,'E:/Matlab_2016a/Hahaha','newhahaha');

function RewriteXLSX(originalfile, sheetname, Nnewfile, target, Ntarget, savedir, newfilename)
[~,~,File0]=xlsread(originalfile,sheetname);
Nrow=size(File0,1);
Ncol=size(File0,2);
cd(savedir);
for i=1:Nnewfile
    switch target
        case 'row'
            Nnew=randsample(1:Nrow,Ntarget);
            File1=File0(Nnew,:);
        case 'col'
            Nnew=randsample(1:Ncol,Ntarget);
            File1=File0(:,Nnew);
    end
    xlswrite([newfilename, num2str(i)],File1)
end
end