%%  XLSX_Splitter.m
%   splits multi-sheet .xlsx files based into mutliple .csv files
%   each sheet becomes its own file
%   files given name of the sheets
%
%   Created 11/13/2021 Dylan Terstege (https://github.com/dterstege)

[file,path] = uigetfile('.xlsx');
sheets = sheetnames(fullfile(path,file));

for ii = 1:numel(sheets)
    M = readcell(fullfile(path,file),'Sheet',sheets(ii));
    outname=fullfile(path,strcat(sheets(ii),'.csv'));
    writecell(M,outname); %write csv
end

clear file path sheets ii M outname