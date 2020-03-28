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
% SheetNames={'A','B','C','D'};
% Nrow=[20 20 20 20];
% RewriteXLSX('hahaha',SheetNames,30,'row',Nrow,'E:/Matlab_2016a/Hahaha','newhahaha');
%

function RewriteXLSX(originalfile, sheetname, Nnewfile, target, Ntarget, savedir, newfilename)
fprintf('Now we start generating new xlsx files...\n')
cd(savedir);
Nsheet=length(sheetname);
for sheet=1:Nsheet
    [~,~,File0]=xlsread(originalfile,sheetname{sheet});
    Nrow=size(File0,1);
    Ncol=size(File0,2);
    for i=1:Nnewfile
        switch target
            case 'row'
                Nnew=randsample(1:Nrow,Ntarget(sheet));
                File1=File0(Nnew,:);
            case 'col'
                Nnew=randsample(1:Ncol,Ntarget(sheet));
                File1=File0(:,Nnew);
        end
        xlswrite([newfilename, num2str(i)],File1,sheetname{sheet})
        fprintf('File: %d, Sheet: %s, finished.\n',i,sheetname{sheet})
    end
end
fprintf('All done.')
end