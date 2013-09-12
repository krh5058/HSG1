[~,d] = system('dir /b/ad');
d = regexp(d(1:end-1),'\n','split');

templatefill = 1;
timingfill = 1;

if templatefill
    [num1,v1] = xlsread('timing_template_v1.xlsx');
    num1 = num2cell(num1);
    num1(cellfun(@isnan,num1)) = {''};
    v1(2:end,1) = num1(:,1);
    v1(2:end,4:end) = num1(:,4:end);
    
    [num2,v2] = xlsread('timing_template_v2.xlsx');
    num2 = num2cell(num2);
    num2(cellfun(@isnan,num2)) = {''};
    v2(2:end,1) = num2(:,1);
    v2(2:end,4:end) = num2(:,4:end);
    
    for i = 1:length(d)
        sid = d{i};
        [~,d2] = system(['dir /b/a-d ' d{i}]);
        d2 = regexp(d2(1:end-1),'\n','split');
        
        for ii = 1:length(d2)
            f_out = regexprep(d2{ii},'raw','');
            [~,d_out] = fileparts(f_out);
            d_fullout = [d{i} filesep d_out];
            mkdir(d_fullout);
            
            strmatch = regexp(d2{ii},'_\d{1,1}','match');
            sess = strmatch{1}(2);
            ver = strmatch{2}(2);
            
            [num,txt] = xlsread([d{i} filesep d2{ii}]);
            num = num2cell(num);
            
            if size(num,2) < 5
                num = [num, num2cell(nan([length(num) (5-size(num,2))]))];
            end
            
            num(cellfun(@isnan,num)) = {''};
            txt(2:end,1) = num(:,1);
            txt(2:end,3:end) = num(:,3:end);
            txt = txt(2:end,:);
            
            % headers: Trial, Trial Type, Dropped, Correct, Incorrect
            
            if strcmp(ver,'1')
                outcell = v1(2:end,:);
                outhead = v1(1,:);
            else
                outcell = v2(2:end,:);
                outhead = v2(1,:);
            end
            
            trialindex = find(~cellfun(@isempty,outcell(:,1)));
            trialref = cell2mat(outcell(trialindex));
            
            dropcol = strcmp(outhead,'Dropped');
            errvgscol = strcmp(outhead,'ErrorVGS');
            erranticol = strcmp(outhead,'ErrorAnti');
            
            for iii = 1:length(txt)
                trialval = txt{iii,1};
                index = trialindex(find(trialref==trialval));
                
                trialtype = txt{iii,2};
                switch trialtype(1)
                    case 'N'
                        horder = 'Neu';
                    case 'R'
                        horder = 'Rew';
                    case 'L'
                        horder = 'Loss';
                end
                
                typeregcheck = regexp(trialtype,'(ANTI)|(VGS)','match');
                switch typeregcheck{1}
                    case 'ANTI'
                        suborder = 'Anti';
                        errcol = erranticol;
                    case 'VGS'
                        suborder = 'Vgs';
                        errcol = errvgscol;
                end
                
                dcival = find(~cellfun(@isempty,txt(iii,3:end)));
                
                switch dcival
                    case 1
                        outcell{index,dropcol} = 1;
                    case 2
                        incol = ~cellfun(@isempty,cellfun(@(y)(regexp(y,[horder suborder])),outhead,'UniformOutput',false));
                        outcell(index:index+2,incol) = num2cell(diag(ones(3,1)));
                    case 3
                        outcell{index,errcol} = 1;
                end
                
            end
            
            xlswrite([d_fullout filesep f_out],[outhead;outcell]);
            
            if timingfill
                subcell = {'anti','vgs'};
                hcell = {'rew','loss','neu'};
                dvcell = {'cue','prep','sac'};
                onsetcol = 5;
                
                for j = 1:length(subcell)
                   for jj = 1:length(hcell)
                      for jjj = 1:length(dvcell)
                          f2_out = [sid '_' hcell{jj} dvcell{jjj} '_' subcell{j} '.1D'];
                          fid = fopen([d_fullout filesep f2_out],'w');
                          
                          colindex = ~cellfun(@isempty,regexpi(outhead,[hcell{jj} subcell{j} dvcell{jjj}]));
                          onsets = outcell(cellfun(@(y)(y==1),outcell(:,colindex)),onsetcol);
                          onsetsformatted = cellfun(@(y)([y ' ']),cellfun(@num2str,onsets,'UniformOutput',false),'UniformOutput',false);
                          
                          fprintf(fid,'%s',[onsetsformatted{:}]);
                          fclose(fid);
                      end
                   end
                end
                
                extracell = {'dropped','errorvgs', 'erroranti'};
                
                for k = 1:length(extracell)
                    f2_out = [sid '_' extracell{k} '.1D'];
                    fid = fopen([d_fullout filesep f2_out],'w');
                    
                    colindex = ~cellfun(@isempty,regexpi(outhead,extracell{k}));
                    onsets = outcell(cellfun(@(y)(y==1),outcell(:,colindex)),onsetcol);
                    onsetsformatted = cellfun(@(y)([y ' ']),cellfun(@num2str,onsets,'UniformOutput',false),'UniformOutput',false);
                    fprintf(fid,'%s',[onsetsformatted{:}]);
                    fclose(fid);
                end
            end
            
        end
        
    end
    
end