classdef StickPicture_exported < matlab.apps.AppBase

    % Properties that correspond to app components
    properties (Access = public)
        UIFigure           matlab.ui.Figure
        Menu               matlab.ui.container.Menu
        OpenMenu           matlab.ui.container.Menu
        QuitMenu           matlab.ui.container.Menu
        ColorMenu          matlab.ui.container.Menu
        Color              matlab.ui.container.Menu
        Grid               matlab.ui.container.Menu
        GroundColor        matlab.ui.container.Menu
        GroundEdge         matlab.ui.container.Menu
        LineMenu           matlab.ui.container.Menu
        xlsxMenu           matlab.ui.container.Menu
        Body23             matlab.ui.container.Menu
        OpenList           matlab.ui.container.Menu
        ButtonGroup        matlab.ui.container.ButtonGroup
        LockButton         matlab.ui.control.RadioButton
        FreeButton         matlab.ui.control.RadioButton
        XZButton           matlab.ui.control.RadioButton
        YZButton           matlab.ui.control.RadioButton
        XYButton           matlab.ui.control.RadioButton
        TimeLabel          matlab.ui.control.Label
        TimeCounter        matlab.ui.control.Label
        LimitrateCheckBox  matlab.ui.control.CheckBox
        Rec                matlab.ui.control.Image
        Pause              matlab.ui.control.Image
        Backward           matlab.ui.control.Image
        Forward            matlab.ui.control.Image
        Stop               matlab.ui.control.Image
        Play               matlab.ui.control.Image
        FrameSpinnerLabel  matlab.ui.control.Label
        FrameSpinner       matlab.ui.control.Spinner
        SpeedKnob          matlab.ui.control.Knob
        SpeedKnobLabel     matlab.ui.control.Label
        FrameSlider        matlab.ui.control.Slider
        UIAxes             matlab.ui.control.UIAxes
    end

    
    properties (Access = public)
        path
        fname
        ifr
        np
        ln
        dot
        p_tbl
        l_tbl
        d_tbl
        nline
        ndot
        pl
        pd
    end
    
    properties (Access = private)
        nf
        dt
        dim
        tt
        xx
        yy
        zz
        ax
        margine
        msh
        speed
        start_frm
        ListApp
    end
    
    methods (Access = private)
        
        function OpenFile(app)
            %% ファイルの読み込み
            [file, path] = uigetfile('*.*');
            if file==0
                return;
            end
            [filepath,fname,ext] = fileparts([path file]);
            app.path = path;
            app.fname = fname;
            header = readmatrix([path file],'FileType','text','Range','A1:C1');
            nf = header(1);
            np = header(2);
            dt = header(3);
            dat = readmatrix([path file],'FileType','text','NumHeaderLines',1);
            dim = size(dat,2)/np;
            tmp_tt = 0:dt:(nf-1)*dt;
            tt = tmp_tt';
            
            %% 設定ファイルの読み込み
            filename = [path fname '.xlsx'];
            result = isfile(filename);
            p_tbl = table();
            l_tbl = table();
            d_tbl = table();
            if result == 1
                % リストファイルがある場合
                tmp_tbl = readtable(filename,'Sheet','point');
                p_tbl.point = tmp_tbl.point;
                p_tbl.name = string(tmp_tbl.name);
                
                tmp_tbl = readtable(filename,'Sheet','line');
                if height(tmp_tbl) > 0
                    l_tbl.list = tmp_tbl.list;
                    l_tbl.line = string(tmp_tbl.line);
                    l_tbl.color = categorical(tmp_tbl.color,{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
                    l_tbl.type = categorical(tmp_tbl.type,{'実線','破線','点線','一点鎖線'});
                    l_tbl.width = tmp_tbl.width;
                end
                
                tmp_tbl = readtable(filename,'Sheet','dot');
                if height(tmp_tbl) > 0
                    d_tbl.list = tmp_tbl.list;
                    d_tbl.name = string(tmp_tbl.name);
                    d_tbl.marker = categorical(tmp_tbl.marker,{'円','プラス記号','アスタリスク','点','十字','正方形','菱形','上向き三角形','下向き三角形','右向き三角形','左向き三角形','星形五角形','星形六角形'});
                    d_tbl.size = tmp_tbl.size;
                    d_tbl.edge = categorical(tmp_tbl.edge,{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
                    d_tbl.face = categorical(tmp_tbl.face,{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
                end
                
                % 変数に格納
                app.p_tbl = p_tbl;
                app.l_tbl = l_tbl;
                app.d_tbl = d_tbl;
                
                [ln,dot] = ConvertLine(app);
                app.ln = ln;
                app.dot = dot;
            else
                % リストファイルがない場合
                for ipnt = 1:np
                    point(ipnt,1) = ipnt;
                    point_name(ipnt,1) = string(ipnt);
                    dot_list(ipnt,1) = ipnt;
                    dot_name(ipnt,1) = string(ipnt);
                    dot_marker(ipnt,1) = "円";
                    dot_size(ipnt,1) = 3;
                    dot_edge(ipnt,1) = "なし";
                    dot_face(ipnt,1) = "白";
                end
                p_tbl.point = point;
                p_tbl.name = point_name;
                d_tbl.list = dot_list;
                d_tbl.name = dot_name;
                d_tbl.marker = dot_marker;
                d_tbl.size = dot_size;
                d_tbl.edge = dot_edge;
                d_tbl.face = dot_face;
                
                % 変数に格納
                app.p_tbl = p_tbl;
                app.l_tbl = l_tbl;
                app.d_tbl = d_tbl;
                
                [ln,dot] = ConvertLine(app);
                app.ln = ln;
                app.dot = dot;
            end
            
            %% x,y,z座標の設定
            for ifr = 1:nf
                for ipnt = 1:np
                    xx(ifr,ipnt) = dat(ifr,(ipnt-1)*dim+1);
                    yy(ifr,ipnt) = dat(ifr,(ipnt-1)*dim+2);
                    if dim == 3
                        zz(ifr,ipnt) = dat(ifr,ipnt*dim);
                    else
                        zz(ifr,ipnt) = 0;
                    end
                end
            end
            min_xx = min(xx,[],'all');
            min_yy = min(yy,[],'all');
            min_zz = min(zz,[],'all');
            max_xx = max(xx,[],'all');
            max_yy = max(yy,[],'all');
            max_zz = max(zz,[],'all');
            margine = 0.05;
            ax = [min_xx-margine,max_xx+margine,min_yy-margine,max_yy+margine,min_zz-margine,max_zz+margine];
                        
            % プライベート変数に格納
            app.nf = nf;
            app.np = np;
            app.dt = dt;
            app.dim = dim;
            app.tt = tt;
            app.xx = xx;
            app.yy = yy;
            app.zz = zz;
            app.ax = ax;
            app.margine = margine;
            
            % 地面にメッシュを作成
            msh = DrawMesh(app);
            
            % プロット初期設定
            ifr = 1;
            x = xx(ifr,:);
            y = yy(ifr,:);
            z = zz(ifr,:);
            
            %% ラインプロット
            %ln = app.ln;
            nline = app.nline;
            if nline > 0
                pl = CreateLine(app,1);
            else
                pl = [];
            end
            
            %% ドットプロット
            ndot = app.ndot;
            if ndot > 0
                pd = CreateDot(app,1);
            else
                pd = [];
            end
            
            %% ビューの設定
            app.UIAxes.Interactions = [rotateInteraction zoomInteraction];  % 回転・ズームのみ操作を許可する
            axis(app.UIAxes, 'equal');
            axis(app.UIAxes,ax);
            if dim == 3
                view(app.UIAxes,[-45 45]);
                app.FreeButton.Value = 1;
                enableDefaultInteractivity(app.UIAxes); % 座標軸の組み込み操作を有効にする
            else
                view(app.UIAxes,[0 90]);
                app.LockButton.Value = 1;
                disableDefaultInteractivity(app.UIAxes);  % 座標軸の組み込み操作を無効にする
            end
            
            %% 変数の設定
            app.pl = pl;
            app.pd = pd;
            app.msh = msh;

            %% ボタンの有効化
            app.Play.Enable = 1;
            app.Stop.Enable = 1;
            app.Backward.Enable = 1;
            app.Forward.Enable = 1;
            app.Rec.Enable = 1;
            app.ButtonGroup.Enable = 'on';
            app.SpeedKnob.Enable = 1;
            app.LimitrateCheckBox.Enable = 1;
            app.ColorMenu.Enable = 1;
            app.LineMenu.Enable = 1;
            
            %% フレームスライダーの設定
            app.FrameSlider.Enable = 1;
            app.FrameSlider.Value = 1;
            app.FrameSlider.Limits = [1,nf];
            
            %% フレームスピナーの設定
            app.FrameSpinner.Enable = 1;
            app.ifr = ifr;
            app.FrameSpinner.Value = ifr;
            app.start_frm = 1;
            
            %% スピードノブの設定
            app.speed = app.SpeedKnob.Value/100;
            
            %% タイムカウンターの設定
            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
            
            %% プロット図にフォーカス
            figure(app.UIFigure);
        end        
    end
    
    methods (Access = public)
        
        function [ln,dot] = ConvertLine(app)
            p_tbl = app.p_tbl;
            l_tbl = app.l_tbl;
            d_tbl = app.d_tbl;
            nline = height(l_tbl);
            ndot = height(d_tbl);
            pname = p_tbl.name;
            app.nline = nline;
            app.ndot = ndot;
            if nline > 0
                for iline = 1:nline
                    % 名称をポイント番号に変換
                    tmp_ln = l_tbl.line(iline);
                    dat_split = split(tmp_ln,'-');
                    nsplit = length(dat_split);
                    p_id = [];
                    for isplit = 1:nsplit
                        p = dat_split(isplit);
                        for irow = 1:height(p_tbl)
                            tmp_name = pname(irow);
                            if p == tmp_name
                                point_no = irow;
                                break;
                            else
                                dummy=1;
                            end
                        end
                        p_id(isplit) = point_no;
                    end
                    ln(iline).p_id = p_id;
                    
                    % 線の色を変換
                    clr = l_tbl.color(iline);
                    
                    if clr == "白"
                        line_color = 'w';
                    elseif clr == "赤"
                        line_color = 'r';
                    elseif clr == "緑"
                        line_color = 'g';
                    elseif clr == "青"
                        line_color = 'b';
                    elseif clr == "黃"
                        line_color = 'y';
                    elseif clr == "マゼンタ"
                        line_color = 'm';
                    elseif clr == "シアン"
                        line_color = 'c';
                    elseif clr == "黒"
                        line_color = 'k';
                    elseif clr == "なし"
                        line_color = 'none';
                    end
                    ln(iline).color = line_color;

                    % 線の種類を変換
                    typ = l_tbl.type(iline);
                    if typ == "実線"
                        line_type = '-';
                    elseif typ == "破線"
                        line_type = '--';
                    elseif typ == "点線"
                        line_type = ':';
                    elseif typ == "一点鎖線"
                        line_type = '-.';
                    end
                    ln(iline).type = line_type;

                    
                    % 線の色を変換
                    line_width = l_tbl.width(iline);
                    ln(iline).width = line_width;
                end
            else
                ln =[];
            end
            
            if ndot > 0
                for idot = 1:ndot
                    % 名称をポイント番号に変換
                    tmp_dot = d_tbl.name(idot);
                    dat_split = split(tmp_dot,'-');
                    nsplit = length(dat_split);
                    for irow = 1:height(p_tbl)
                        tmp_name = pname(irow);
                        if tmp_dot == tmp_name
                            point_no = irow;
                            break;
                        else
                            dummy=1;
                        end
                    end
                    dot(idot).point = point_no;
                    
                    mkr = d_tbl.marker(idot);
                    if mkr == "円"
                        marker = 'o';
                    elseif mkr == "プラス記号"
                        marker = '+';
                    elseif mkr == "アスタリスク"
                        marker = '*';
                    elseif mkr == "点"
                        marker = '.';
                    elseif mkr == "十字"
                        marker = 'x';
                    elseif mkr == "正方形"
                        marker = 's';
                    elseif mkr == "菱形"
                        marker = 'd';
                    elseif mkr == "上向き三角形"
                        marker = '^';
                    elseif mkr == "下向き三角形"
                        marker = 'v';
                    elseif mkr == "右向き三角形"
                        marker = '>';
                    elseif mkr == "左向き三角形"
                        marker = '<';
                    elseif mkr == "星形五角形"
                        marker = 'p';
                    elseif mkr == "星形六角形"
                        marker = 'h';
                    end
                    dot(idot).marker = marker;
                    
                    dot(idot).size = d_tbl.size(idot);
                    
                    clr = d_tbl.edge(idot);
                    if clr == "なし"
                        edge_color = 'none';
                    elseif clr == "白"
                        edge_color = 'w';
                    elseif clr == "赤"
                        edge_color = 'r';
                    elseif clr == "緑"
                        edge_color = 'g';
                    elseif clr == "青"
                        edge_color = 'b';
                    elseif clr == "黃"
                        edge_color = 'y';
                    elseif clr == "マゼンタ"
                        edge_color = 'm';
                    elseif clr == "シアン"
                        edge_color = 'c';
                    elseif clr == "黒"
                        edge_color = 'k';
                    end
                    dot(idot).edge = edge_color;
                    
                    clr = d_tbl.face(idot);
                    if clr == "なし"
                        face_color = 'none';
                    elseif clr == "白"
                        face_color = 'w';
                    elseif clr == "赤"
                        face_color = 'r';
                    elseif clr == "緑"
                        face_color = 'g';
                    elseif clr == "青"
                        face_color = 'b';
                    elseif clr == "黃"
                        face_color = 'y';
                    elseif clr == "マゼンタ"
                        face_color = 'm';
                    elseif clr == "シアン"
                        face_color = 'c';
                    elseif clr == "黒"
                        face_color = 'k';
                    end
                    dot(idot).face = face_color;
                end
            else
                dot = [];
            end
        end
        
        function pl = CreateLine(app,ifr)
            nline = app.nline;
            ln = app.ln;
            x = app.xx(ifr,:);
            y = app.yy(ifr,:);
            z = app.zz(ifr,:);
            for iline = 1:nline
                p_id = ln(iline).p_id;
                line_type = ln(iline).type;
                line_color = ln(iline).color;
                line_width = ln(iline).width;
                x1 = [];
                y1 = [];
                z1 = [];
                for id = 1:length(p_id)
                    x1(id) = x(p_id(id));
                    y1(id) = y(p_id(id));
                    z1(id) = z(p_id(id));
                end
                pl(iline) = line(app.UIAxes,x1,y1,z1);
                pl(iline).Color = line_color;
                pl(iline).LineStyle = line_type;
                pl(iline).LineWidth = line_width;
            end
        end
        
        function pd = CreateDot(app,ifr)
            ndot= app.ndot;
            dot = app.dot;
            x = app.xx(ifr,:);
            y = app.yy(ifr,:);
            z = app.zz(ifr,:);
            for idot = 1:ndot
                point = dot(idot).point;
                marker = dot(idot).marker;
                dot_size = dot(idot).size;
                dot_edge = dot(idot).edge;
                dot_face = dot(idot).face;
                x1 = x(point);
                y1 = y(point);
                z1 = z(point);
                pd(idot) = line(app.UIAxes,x1,y1,z1);
                pd(idot).Marker = marker;
                pd(idot).MarkerSize = dot_size;
                pd(idot).MarkerEdgeColor = dot_edge;
                pd(idot).MarkerFaceColor = dot_face;
            end
        end
        
        function PlotLine(app,ifr)
            pl = app.pl;
            ln = app.ln;
            nline = app.nline;
            x = app.xx(ifr,:);
            y = app.yy(ifr,:);
            z = app.zz(ifr,:);
            for iline = 1:nline
                p_id = ln(iline).p_id;
                x1 = [];
                y1 = [];
                z1 = [];
                for id = 1:length(p_id)
                    x1(id) = x(p_id(id));
                    y1(id) = y(p_id(id));
                    z1(id) = z(p_id(id));
                end
                set(pl(iline),'XData',x1,'YData',y1,'ZData',z1);
            end
        end
        
        function PlotDot(app,ifr)
            pd = app.pd;
            dot = app.dot;
            ndot = app.ndot;
            x = app.xx(ifr,:);
            y = app.yy(ifr,:);
            z = app.zz(ifr,:);
            for idot = 1:ndot
                point = dot(idot).point;
                set(pd(idot),'XData',x(point),'YData',y(point),'ZData',z(point));
            end
        end
        
        function msh = DrawMesh(app)
            ax = app.ax;
            dim = app.dim;
            margine = app.margine;
        	minAx = ax(1) + margine;
            maxAx = ax(2) - margine;
            minAy = ax(3) + margine;
            maxAy = ax(4) - margine;
        
        	MeshSize=0.5;   % メッシュの大きさ（default:0.5）
        	XX = [minAx:MeshSize:maxAx];
        	YY = [minAy:MeshSize:maxAy];
        	LX = length(XX);
        	LY = length(YY);
        	if (LX > LY)
        		dL = LX - LY;
        		for j = 1:dL
        			YY = [YY,maxAy];
        		end
        	else
        		dL = LY - LX;
        		for j = 1:dL
        			XX = [XX,maxAx];
        		end
        	end
        	XX = [XX,maxAx];
        	YY = [YY,maxAy];
        	scl = length(XX);
        	ZZ = zeros(scl,scl);    % z軸(高さ)はゼロ(床面)
        	c = [0 0 0.7]';         % 床面の色([R G B]')
        	if dim == 2
        		YY = ZZ;
        	end
        	msh = mesh(app.UIAxes,XX,YY,ZZ,'EdgeColor',c,'FaceColor','k');   % 3次元の場合は床にメッシュを入れる
        end
        
    end
        

    % Callbacks that handle component events
    methods (Access = private)

        % Menu selected function: OpenMenu
        function OpenMenuSelected(app, event)
            OpenFile(app);
        end

        % Image clicked function: Forward
        function ForwardClicked(app, event)
            ifr = app.FrameSpinner.Value;
            nf = app.nf;
            if ifr + 1 > nf
                ifr = nf;
            else
                ifr = ifr + 1;
            end
            app.FrameSpinner.Value = ifr;
            app.FrameSlider.Value = ifr;
            tt = app.tt;
            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
            PlotLine(app,ifr);
            PlotDot(app,ifr);
            drownow;
        end

        % Image clicked function: Backward
        function BackwardImageClicked(app, event)
            ifr = app.FrameSpinner.Value;
            if ifr - 1 == 0
                ifr = 1;
            else
                ifr = ifr - 1;
            end
            app.FrameSpinner.Value = ifr;
            app.FrameSlider.Value = ifr;
            tt = app.tt;
            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
            PlotLine(app,ifr);
            PlotDot(app,ifr);
            drawnow;
        end

        % Value changed function: FrameSlider
        function FrameSliderValueChanged(app, event)
            value = app.FrameSlider.Value;
            
        end

        % Value changing function: FrameSlider
        function FrameSliderValueChanging(app, event)
            if app.Play.Visible == 0
                % 再生中の場合、一時停止する
                app.Pause.Visible = 0;
            end
            changingValue = event.Value;
            ifr = round(changingValue);
            app.FrameSpinner.Value = ifr;
            tt = app.tt;
            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
            PlotLine(app,ifr);
            PlotDot(app,ifr);
            drawnow;
        end

        % Image clicked function: Play
        function PlayImageClicked(app, event)
            app.Pause.Visible = 1;
            app.Play.Visible = 0;
            nf = app.nf;
            dt = app.dt;
            tt = app.tt;
            start_frm = app.FrameSpinner.Value;
            ifr = app.start_frm;
            tStart = tic;
            while app.Pause.Visible == 1
                %start_frm = app.start_frm;
                speed = app.speed;
                ifr = start_frm + round(toc(tStart)*speed/dt);
                if ifr > nf
                    ifr = 1;
                    start_frm = 1;
                    tStart = tic;
                end
                app.FrameSpinner.Value = ifr;
                app.FrameSlider.Value = ifr;
                app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
                PlotLine(app,ifr);
                PlotDot(app,ifr);
                if app.LimitrateCheckBox.Value == 1
                    drawnow limitrate;
                else
                    drawnow;
                end
            end
            app.Play.Visible = 1;
        end

        % Image clicked function: Pause
        function PauseImageClicked(app, event)
            app.Pause.Visible = 0;
            ifr = app.FrameSpinner;
        end

        % Value changing function: SpeedKnob
        function SpeedKnobValueChanging(app, event)
            speed = event.Value/100;
            app.speed = speed;
            app.SpeedKnobLabel.Text = ["Speed" string(round(speed*100,1))];
            %app.start_frm = app.FrameSpinner.Value
            tic;
        end

        % Value changed function: SpeedKnob
        function SpeedKnobValueChanged(app, event)
            ifr = app.FrameSpinner.Value;
            speed = app.SpeedKnob.Value/100;
            app.speed = speed;
            app.SpeedKnobLabel.Text = ["Speed" string(round(speed*100,1))];
            %app.start_frm = app.FrameSpinner.Value;
            tic;
        end

        % Image clicked function: Stop
        function StopImageClicked(app, event)
            if app.Play.Visible == 0
                % 再生中の場合、一時停止する
                app.Pause.Visible = 0;
            end
            ifr = 1;
            app.FrameSpinner.Value = ifr;
            app.FrameSlider.Value = ifr;
            tt = app.tt;
            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
            PlotLine(app,ifr);
            PlotDot(app,ifr);
            drawnow;
        end

        % Callback function
        function OpenListSelected(app, event)
%             app.ListApp = lst(app);
%             app.ln;
%             list = app.list
%             ln = app.ln
        end

        % Image clicked function: Rec
        function RecImageClicked(app, event)
            nf = app.nf;
            dt = app.dt;
            tt = app.tt;
            fps = round(1/dt);

            avibutton = questdlg('アニメーションビデオを作成しますか?','動画ファイルの作成');
        	if strcmp(avibutton,'Yes')
        		comp = 'None';
        		loop = 1;
        	elseif strcmp(avibutton,'No')
        		return;
        	elseif strcmp(avibutton,'Cancel')
        		return;
            end
            [moviefile,moviepath] = uiputfile('*.mp4','Save movie file');
        	if moviefile == 0
        	  return;
            end
        	Pth = [moviepath moviefile];
            
            % メニューの非表示化
            app.Menu.Visible = 0;
            app.ColorMenu.Visible = 0;
            
        	% Movieファイルの初期設定
        	v = VideoWriter(Pth, 'MPEG-4');
        	v.Quality = 100;
        	if fps > 172
        		n_fps = 172;
        	else
        		n_fps = fps;
        	end
        	v.FrameRate = n_fps;	% MAX:172
        	open(v);
         
            for ifr = 1:nf
                PlotLine(app,ifr);
                PlotDot(app,ifr);
                drawnow limitrate;
                tmp_frame = getframe(app.UIAxes);
                writeVideo(v,tmp_frame);
            end
            
            % メニューの再表示
            app.Menu.Visible = 1;
            app.ColorMenu.Visible = 1;
            
            close(v);
        end

        % Menu selected function: Color
        function ColorMenuSelected(app, event)
            c = uisetcolor;
            if length(c) == 1
                return;
            end
            app.UIAxes.Color = c;
            figure(app.UIFigure);
        end

        % Menu selected function: Grid
        function GridMenuSelected(app, event)
            c = uisetcolor;
            if length(c) == 1
                return;
            end
            set(app.UIAxes,'XColor',c,'YColor',c,'ZColor',c);
            figure(app.UIFigure);
        end

        % Menu selected function: GroundColor
        function GroundColorMenuSelected(app, event)
            c = uisetcolor;
            if length(c) == 1
                return;
            end
            set(app.msh,'FaceColor',c);
            figure(app.UIFigure);
        end

        % Menu selected function: GroundEdge
        function GroundEdgeMenuSelected(app, event)
            c = uisetcolor;
            if length(c) == 1
                return;
            end
            set(app.msh,'EdgeColor',c);
            figure(app.UIFigure);
        end

        % Selection changed function: ButtonGroup
        function ButtonGroupSelectionChanged(app, event)
            %selectedButton = app.ButtonGroup.SelectedObject;
            if app.FreeButton.Value == 1
                enableDefaultInteractivity(app.UIAxes); % 座標軸の組み込み操作を有効にする
            elseif app.LockButton.Value == 1
                disableDefaultInteractivity(app.UIAxes);  % 座標軸の組み込み操作を無効にする
            elseif app.XYButton.Value == 1
                view(app.UIAxes,[0 90]);
            elseif app.XZButton.Value == 1
                view(app.UIAxes,[0 0]);
            elseif app.YZButton.Value == 1
                view(app.UIAxes,[90 0]);
            end
        end

        % Menu selected function: OpenList
        function OpenListMenuSelected(app, event)
            app.ListApp = lst(app);
        end

        % Menu selected function: xlsxMenu
        function xlsxMenuSelected(app, event)
            ifr = app.FrameSpinner.Value;
            %% ファイルの読み込み
            [file, path] = uigetfile('*.xlsx');
            if file==0
                return;
            end
            
            %% 設定ファイルの読み込み
            filename = [path file];

            %% ラインの初期化
            for iline = 1:app.nline
                set(app.pl(iline),'XData',[],'YData',[],'ZData',[]);
            end
            
            %% ドットの初期化
            pd = app.pd;
            for idot = 1:app.ndot
                set(app.pd(idot),'XData',[],'YData',[],'ZData',[]);
            end
            
            %% テーブルの初期化
            p_tbl = table();
            l_tbl = table();
            d_tbl = table();
            app.pl = [];
            app.pd = [];
            
            tmp_tbl = readtable(filename,'Sheet','point');
            p_tbl.point = tmp_tbl.point;
            p_tbl.name = string(tmp_tbl.name);
            
            tmp_tbl = readtable(filename,'Sheet','line');
            if height(tmp_tbl) > 0
                l_tbl.list = tmp_tbl.list;
                l_tbl.line = string(tmp_tbl.line);
                l_tbl.color = categorical(tmp_tbl.color,{'白','赤','緑','青','黄','マゼンタ','シアン','黒','なし'});
                l_tbl.type = categorical(tmp_tbl.type,{'実線','破線','点線','一点鎖線'});
                l_tbl.width = tmp_tbl.width;
            end
            
            tmp_tbl = readtable(filename,'Sheet','dot');
            if height(tmp_tbl) > 0
                d_tbl.list = tmp_tbl.list;
                d_tbl.name = string(tmp_tbl.name);
                d_tbl.marker = categorical(tmp_tbl.marker,{'円','プラス記号','アスタリスク','点','十字','正方形','菱形','上向き三角形','下向き三角形','右向き三角形','左向き三角形','星形五角形','星形六角形'});
                d_tbl.size = tmp_tbl.size;
                d_tbl.edge = categorical(tmp_tbl.edge,{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
                d_tbl.face = categorical(tmp_tbl.face,{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
            end
            
            % 変数に格納
            app.p_tbl = p_tbl;
            app.l_tbl = l_tbl;
            app.d_tbl = d_tbl;
            
            [ln,dot] = ConvertLine(app);
            app.ln = ln;
            app.dot = dot;
            
            %% ラインプロット
            %ln = app.ln;
            nline = app.nline;
            if nline > 0
                pl = CreateLine(app,ifr);
            else
                pl = [];
            end
            
            %% ドットプロット
            ndot = app.ndot;
            if ndot > 0
                pd = CreateDot(app,ifr);
            else
                pd = [];
            end
            
            app.pl = pl;
            app.pd = pd;
            
            drawnow;
            figure(app.UIFigure);
        end

        % Menu selected function: Body23
        function Body23MenuSelected(app, event)
            ifr = app.FrameSpinner.Value;
            %np = 23;
            np = app.np;

            %% ラインの初期化
            for iline = 1:app.nline
                set(app.pl(iline),'XData',[],'YData',[],'ZData',[]);
            end
            
            %% ドットの初期化
            for idot = 1:app.ndot
                set(app.pd(idot),'XData',[],'YData',[],'ZData',[]);
            end
            
            %% テーブルの初期化
            p_tbl = table();
            l_tbl = table();
            d_tbl = table();
            app.pl = [];
            app.pd = [];
            pname = ["右手","右手首","右肘","右肩",...
                     "左手","左手首","左肘","左肩",...
                     "右つま先","右拇指球","右踵","右足","右膝","右大転子",...
                     "左つま先","左拇指球","左踵","左足","左膝","左大転子",...
                     "頭頂","耳珠点","胸骨上縁"];
            
            for ipnt = 1:np
                point(ipnt,1) = ipnt;
                if ipnt <= 23
                    point_name(ipnt,1) = pname(ipnt);
                else
                    point_name(ipnt,1) = string(ipnt);
                end
            end
            p_tbl.point = point;
            p_tbl.name = point_name;
            
            body_line = ["右手-右手首-右肘-右肩-左肩",...
                         "左手-左手首-左肘-左肩",...
                         "右つま先-右拇指球-右踵-右足-右膝-右大転子",...
                         "左つま先-左拇指球-左踵-左足-左膝-左大転子",...
                         "頭頂-耳珠点-胸骨上縁",...
                         "右肩-右大転子",...
                         "左肩-左大転子",...
                         "右大転子-左大転子"];
            
            for iline = 1:8
                line_list(iline,1) = iline;
                line_line(iline,1) = body_line(iline);
                switch iline
                    case 1
                        line_color(iline,1) = categorical("白",{'白','赤','緑','青','黄','マゼンタ','シアン','黒','なし'});
                        line_type(iline,1) = categorical("実線",{'実線','破線','点線','一点鎖線'});
                        line_width(iline,1) = 0.5;
                    case 2
                        line_color(iline,1) = "赤";
                        line_type(iline,1) = "破線";
                        line_width(iline,1) = 0.5;
                    case 3
                        line_color(iline,1) = "白";
                        line_type(iline,1) = "実線";
                        line_width(iline,1) = 0.5;
                    case 4
                        line_color(iline,1) = "赤";
                        line_type(iline,1) = "破線";
                        line_width(iline,1) = 0.5;
                    case 5
                        line_color(iline,1) = "白";
                        line_type(iline,1) = "実線";
                        line_width(iline,1) = 0.5; 
                    case 6
                        line_color(iline,1) = "白";
                        line_type(iline,1) = "実線";
                        line_width(iline,1) = 0.5;
                    case 7
                        line_color(iline,1) = "赤";
                        line_type(iline,1) = "破線";
                        line_width(iline,1) = 0.5;
                    case 8
                        line_color(iline,1) = "白";
                        line_type(iline,1) = "実線";
                        line_width(iline,1) = 0.5;
                end
            end
            l_tbl.list = line_list;
            l_tbl.line = line_line;
            l_tbl.color = line_color;
            l_tbl.type = line_type;
            l_tbl.width = line_width;
            
            % 身体23点以降のポイントの設定
            if np > 23
                for idot = 1:np-23
                    dot_list(idot,1) = idot;
                    dot_name(idot,1) = string(idot+23);
                    dot_marker(idot,1) = categorical("円",{'円','プラス記号','アスタリスク','点','十字','正方形','菱形','上向き三角形','下向き三角形','右向き三角形','左向き三角形','星形五角形','星形六角形'});
                    dot_size(idot,1) = 3;
                    dot_edge(idot,1) = categorical("白",{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
                    dot_face(idot,1) = categorical("白",{'なし','白','赤','緑','青','黄','マゼンタ','シアン','黒'});
                end
                d_tbl.list = dot_list;
                d_tbl.name = dot_name;
                d_tbl.marker = dot_marker;
                d_tbl.size = dot_size;
                d_tbl.edge = dot_edge;
                d_tbl.face = dot_face;
            end
            
            % 変数に格納
            app.p_tbl = p_tbl;
            app.l_tbl = l_tbl;
            app.d_tbl = d_tbl;
            
            [ln,dot] = ConvertLine(app);
            app.ln = ln;
            app.dot = dot;
            
            %% ラインプロット
            %ln = app.ln;
            nline = app.nline;
            if nline > 0
                pl = CreateLine(app,ifr);
            else
                pl = [];
            end
            
            %% ドットプロット
            ndot = app.ndot;
            if ndot > 0
                pd = CreateDot(app,ifr);
            else
                pd = [];
            end
            
            app.pl = pl;
            app.pd = pd;
            
            drawnow;
        end

        % Menu selected function: QuitMenu
        function QuitMenuSelected(app, event)
            delete(app);
        end

        % Value changing function: FrameSpinner
        function FrameSpinnerValueChanging(app, event)
            if app.Play.Visible == 0
                % 再生中の場合、一時停止する
                app.Pause.Visible = 0;
            end
            changingValue = event.Value;
            ifr = round(changingValue);
            app.FrameSpinner.Value = ifr;
            tt = app.tt;
            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
            PlotLine(app,ifr);
            PlotDot(app,ifr);
            drawnow;
        end

        % Key press function: UIFigure
        function UIFigureKeyPress(app, event)
            key = event.Key;
            switch key
                case 'rightarrow'
                    ifr = app.FrameSpinner.Value;
                    nf = app.nf;
                    if ifr + 1 > nf
                        ifr = nf;
                    else
                        ifr = ifr + 1;
                    end
                    app.FrameSpinner.Value = ifr;
                    app.FrameSlider.Value = ifr;
                    tt = app.tt;
                    app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
                    PlotLine(app,ifr);
                    PlotDot(app,ifr);
                case 'uparrow'
                    ifr = app.FrameSpinner.Value;
                    nf = app.nf;
                    if ifr + 1 > nf
                        ifr = nf;
                    else
                        ifr = ifr + 1;
                    end
                    app.FrameSpinner.Value = ifr;
                    app.FrameSlider.Value = ifr;
                    tt = app.tt;
                    app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
                    PlotLine(app,ifr);
                    PlotDot(app,ifr);
                case 'leftarrow'
                    ifr = app.FrameSpinner.Value;
                    if ifr - 1 == 0
                        ifr = 1;
                    else
                        ifr = ifr - 1;
                    end
                    app.FrameSpinner.Value = ifr;
                    app.FrameSlider.Value = ifr;
                    tt = app.tt;
                    app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
                    PlotLine(app,ifr);
                    PlotDot(app,ifr);
                case 'downarrow'
                    ifr = app.FrameSpinner.Value;
                    if ifr - 1 == 0
                        ifr = 1;
                    else
                        ifr = ifr - 1;
                    end
                    app.FrameSpinner.Value = ifr;
                    app.FrameSlider.Value = ifr;
                    tt = app.tt;
                    app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
                    PlotLine(app,ifr);
                    PlotDot(app,ifr);
                case 'space'
                    if app.Pause.Visible == 0
                        % 一時停止中は再生
                        app.Pause.Visible = 1;
                        app.Play.Visible = 0;
                        nf = app.nf;
                        dt = app.dt;
                        tt = app.tt;
                        start_frm = app.FrameSpinner.Value;
                        ifr = app.start_frm;
                        tStart = tic;
                        while app.Pause.Visible == 1
                            %start_frm = app.start_frm;
                            speed = app.speed;
                            ifr = start_frm + round(toc(tStart)*speed/dt);
                            if ifr > nf
                                ifr = 1;
                                start_frm = 1;
                                tStart = tic;
                            end
                            app.FrameSpinner.Value = ifr;
                            app.FrameSlider.Value = ifr;
                            app.TimeCounter.Text = sprintf('%.3f',tt(ifr));
                            PlotLine(app,ifr);
                            PlotDot(app,ifr);
                            if app.LimitrateCheckBox.Value == 1
                                drawnow limitrate;
                            else
                                drawnow;
                            end
                        end
                        app.Play.Visible = 1;
                    else
                        % 再生中は一時停止
                        app.Pause.Visible = 0;
                        ifr = app.FrameSpinner;
                    end
            end
        end
    end

    % Component initialization
    methods (Access = private)

        % Create UIFigure and components
        function createComponents(app)

            % Create UIFigure and hide until all components are created
            app.UIFigure = uifigure('Visible', 'off');
            app.UIFigure.Position = [100 100 984 665];
            app.UIFigure.Name = 'MATLAB App';
            app.UIFigure.KeyPressFcn = createCallbackFcn(app, @UIFigureKeyPress, true);

            % Create Menu
            app.Menu = uimenu(app.UIFigure);
            app.Menu.Text = 'ファイル';

            % Create OpenMenu
            app.OpenMenu = uimenu(app.Menu);
            app.OpenMenu.MenuSelectedFcn = createCallbackFcn(app, @OpenMenuSelected, true);
            app.OpenMenu.Text = '座標データを開く';

            % Create QuitMenu
            app.QuitMenu = uimenu(app.Menu);
            app.QuitMenu.MenuSelectedFcn = createCallbackFcn(app, @QuitMenuSelected, true);
            app.QuitMenu.Text = 'アプリ終了';

            % Create ColorMenu
            app.ColorMenu = uimenu(app.UIFigure);
            app.ColorMenu.Enable = 'off';
            app.ColorMenu.Text = '色の変更';

            % Create Color
            app.Color = uimenu(app.ColorMenu);
            app.Color.MenuSelectedFcn = createCallbackFcn(app, @ColorMenuSelected, true);
            app.Color.Text = '背景';

            % Create Grid
            app.Grid = uimenu(app.ColorMenu);
            app.Grid.MenuSelectedFcn = createCallbackFcn(app, @GridMenuSelected, true);
            app.Grid.Text = '軸';

            % Create GroundColor
            app.GroundColor = uimenu(app.ColorMenu);
            app.GroundColor.MenuSelectedFcn = createCallbackFcn(app, @GroundColorMenuSelected, true);
            app.GroundColor.Text = '地面';

            % Create GroundEdge
            app.GroundEdge = uimenu(app.ColorMenu);
            app.GroundEdge.MenuSelectedFcn = createCallbackFcn(app, @GroundEdgeMenuSelected, true);
            app.GroundEdge.Text = '地面枠';

            % Create LineMenu
            app.LineMenu = uimenu(app.UIFigure);
            app.LineMenu.Enable = 'off';
            app.LineMenu.Text = '結線';

            % Create xlsxMenu
            app.xlsxMenu = uimenu(app.LineMenu);
            app.xlsxMenu.MenuSelectedFcn = createCallbackFcn(app, @xlsxMenuSelected, true);
            app.xlsxMenu.Text = 'xlsxファイル読み込み';

            % Create Body23
            app.Body23 = uimenu(app.LineMenu);
            app.Body23.MenuSelectedFcn = createCallbackFcn(app, @Body23MenuSelected, true);
            app.Body23.Text = '身体23点';

            % Create OpenList
            app.OpenList = uimenu(app.LineMenu);
            app.OpenList.MenuSelectedFcn = createCallbackFcn(app, @OpenListMenuSelected, true);
            app.OpenList.Text = '設定';

            % Create UIAxes
            app.UIAxes = uiaxes(app.UIFigure);
            xlabel(app.UIAxes, 'X [m]')
            ylabel(app.UIAxes, 'Y [m]')
            zlabel(app.UIAxes, 'Z [m]')
            app.UIAxes.FontName = 'Times New Roman';
            app.UIAxes.MinorGridLineStyle = '-';
            app.UIAxes.XColor = [0 0 0];
            app.UIAxes.YColor = [0 0 0];
            app.UIAxes.ZColor = [0 0 0];
            app.UIAxes.Color = [0 0 0];
            app.UIAxes.XGrid = 'on';
            app.UIAxes.YGrid = 'on';
            app.UIAxes.ZGrid = 'on';
            app.UIAxes.GridColor = [1 1 1];
            app.UIAxes.MinorGridColor = [1 1 1];
            app.UIAxes.Clipping = 'off';
            app.UIAxes.Box = 'on';
            app.UIAxes.Position = [1 126 960 540];

            % Create FrameSlider
            app.FrameSlider = uislider(app.UIFigure);
            app.FrameSlider.ValueChangedFcn = createCallbackFcn(app, @FrameSliderValueChanged, true);
            app.FrameSlider.ValueChangingFcn = createCallbackFcn(app, @FrameSliderValueChanging, true);
            app.FrameSlider.Enable = 'off';
            app.FrameSlider.Position = [21 123 940 3];
            app.FrameSlider.Value = 1;

            % Create SpeedKnobLabel
            app.SpeedKnobLabel = uilabel(app.UIFigure);
            app.SpeedKnobLabel.HorizontalAlignment = 'center';
            app.SpeedKnobLabel.Enable = 'off';
            app.SpeedKnobLabel.Position = [528 25 40 28];
            app.SpeedKnobLabel.Text = {'Speed'; '100'};

            % Create SpeedKnob
            app.SpeedKnob = uiknob(app.UIFigure, 'continuous');
            app.SpeedKnob.Limits = [0 200];
            app.SpeedKnob.MajorTicks = [0 100 200];
            app.SpeedKnob.ValueChangedFcn = createCallbackFcn(app, @SpeedKnobValueChanged, true);
            app.SpeedKnob.ValueChangingFcn = createCallbackFcn(app, @SpeedKnobValueChanging, true);
            app.SpeedKnob.Enable = 'off';
            app.SpeedKnob.Position = [581 26 49 49];
            app.SpeedKnob.Value = 100;

            % Create FrameSpinner
            app.FrameSpinner = uispinner(app.UIFigure);
            app.FrameSpinner.ValueChangingFcn = createCallbackFcn(app, @FrameSpinnerValueChanging, true);
            app.FrameSpinner.Interruptible = 'off';
            app.FrameSpinner.Enable = 'off';
            app.FrameSpinner.Position = [418 29 89 26];

            % Create FrameSpinnerLabel
            app.FrameSpinnerLabel = uilabel(app.UIFigure);
            app.FrameSpinnerLabel.HorizontalAlignment = 'right';
            app.FrameSpinnerLabel.Position = [371 31 40 22];
            app.FrameSpinnerLabel.Text = 'Frame';

            % Create Play
            app.Play = uiimage(app.UIFigure);
            app.Play.ImageClickedFcn = createCallbackFcn(app, @PlayImageClicked, true);
            app.Play.Enable = 'off';
            app.Play.Tooltip = {'再生'};
            app.Play.Position = [48 25 54 51];
            app.Play.ImageSource = 'play.png';

            % Create Stop
            app.Stop = uiimage(app.UIFigure);
            app.Stop.ImageClickedFcn = createCallbackFcn(app, @StopImageClicked, true);
            app.Stop.Enable = 'off';
            app.Stop.Tooltip = {'停止'};
            app.Stop.Position = [239 24 53 53];
            app.Stop.ImageSource = 'stop.png';

            % Create Forward
            app.Forward = uiimage(app.UIFigure);
            app.Forward.ImageClickedFcn = createCallbackFcn(app, @ForwardClicked, true);
            app.Forward.Enable = 'off';
            app.Forward.Tooltip = {'コマ送り'};
            app.Forward.Position = [171 24 63 53];
            app.Forward.ImageSource = 'forward.png';

            % Create Backward
            app.Backward = uiimage(app.UIFigure);
            app.Backward.ImageClickedFcn = createCallbackFcn(app, @BackwardImageClicked, true);
            app.Backward.Enable = 'off';
            app.Backward.Tooltip = {'コマ戻し'};
            app.Backward.Position = [110 24 63 53];
            app.Backward.ImageSource = 'backward.png';

            % Create Pause
            app.Pause = uiimage(app.UIFigure);
            app.Pause.ImageClickedFcn = createCallbackFcn(app, @PauseImageClicked, true);
            app.Pause.Visible = 'off';
            app.Pause.Tooltip = {'一時停止'};
            app.Pause.Position = [48 24 53 53];
            app.Pause.ImageSource = 'pause.png';

            % Create Rec
            app.Rec = uiimage(app.UIFigure);
            app.Rec.ImageClickedFcn = createCallbackFcn(app, @RecImageClicked, true);
            app.Rec.Enable = 'off';
            app.Rec.Tooltip = {'動画ファイル作成'};
            app.Rec.Position = [311 24 40 53];
            app.Rec.ImageSource = 'rec.png';

            % Create LimitrateCheckBox
            app.LimitrateCheckBox = uicheckbox(app.UIFigure);
            app.LimitrateCheckBox.Enable = 'off';
            app.LimitrateCheckBox.Text = 'Limitrateモードで再生';
            app.LimitrateCheckBox.Position = [373 -6 140 41];

            % Create TimeCounter
            app.TimeCounter = uilabel(app.UIFigure);
            app.TimeCounter.Position = [418 56 63 37];
            app.TimeCounter.Text = 'sec';

            % Create TimeLabel
            app.TimeLabel = uilabel(app.UIFigure);
            app.TimeLabel.HorizontalAlignment = 'center';
            app.TimeLabel.Position = [382 56 26 37];
            app.TimeLabel.Text = 'sec';

            % Create ButtonGroup
            app.ButtonGroup = uibuttongroup(app.UIFigure);
            app.ButtonGroup.SelectionChangedFcn = createCallbackFcn(app, @ButtonGroupSelectionChanged, true);
            app.ButtonGroup.Enable = 'off';
            app.ButtonGroup.TitlePosition = 'centertop';
            app.ButtonGroup.Title = '視点操作';
            app.ButtonGroup.Position = [691 8 142 87];

            % Create XYButton
            app.XYButton = uiradiobutton(app.ButtonGroup);
            app.XYButton.Text = 'X-Y軸';
            app.XYButton.Position = [75 40 58 22];

            % Create YZButton
            app.YZButton = uiradiobutton(app.ButtonGroup);
            app.YZButton.Text = 'Y-Z軸';
            app.YZButton.Position = [75 2 58 22];

            % Create XZButton
            app.XZButton = uiradiobutton(app.ButtonGroup);
            app.XZButton.Text = 'X-Z軸';
            app.XZButton.Position = [75 21 58 22];

            % Create FreeButton
            app.FreeButton = uiradiobutton(app.ButtonGroup);
            app.FreeButton.Text = 'フリー';
            app.FreeButton.Position = [13 36 58 22];
            app.FreeButton.Value = true;

            % Create LockButton
            app.LockButton = uiradiobutton(app.ButtonGroup);
            app.LockButton.Text = '固定';
            app.LockButton.Position = [13 10 58 22];

            % Show the figure after all components are created
            app.UIFigure.Visible = 'on';
        end
    end

    % App creation and deletion
    methods (Access = public)

        % Construct app
        function app = StickPicture_exported

            % Create UIFigure and components
            createComponents(app)

            % Register the app with App Designer
            registerApp(app, app.UIFigure)

            if nargout == 0
                clear app
            end
        end

        % Code that executes before app deletion
        function delete(app)

            % Delete UIFigure when app is deleted
            delete(app.UIFigure)
        end
    end
end