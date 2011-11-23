/*
 * Timeglider for Javascript / jQuery 
 * http://timeglider.com/jquery
 *
 * Copyright 2011, Mnemograph LLC
 * Licensed under Timeglider Dual License
 * http://timeglider.com/jquery/?p=license
 *
/*

*         DEPENDENCIES: 
                        rafael.js
                        ba-tinyPubSub
                        jquery
                        jquery ui (and css)
                        jquery.mousewheel
                        jquery.ui.ipad
                        
                        TG_Date.js
                        TG_Timeline.js
                        TG_TimelineView.js
                        TG_Mediator.js
                        TG_Org.js
                        Timeglider.css
*
*/




(function($){
	/**
	* The main jQuery widget factory for Timeglider
	*/
	
	var timelinePlayer, 
		MED, 
		tg = timeglider, 
		TG_Date = timeglider.TG_Date;
	
	$.widget( "timeglider.timeline", {
		
		// defaults!
		options : { 
			timezone:"00:00",
			initial_focus:tg.TG_Date.getToday(), 
			editor:'none', 
			min_zoom : 1, 
			max_zoom : 100, 
			show_centerline: true, 
			data_source:"", 
			culture:"en",
			basic_fontsize:12, 
			mouse_wheel: "zoom", 
			initial_timeline_id:'',
			icon_folder:'js/timeglider/icons/',
			show_footer:true,
			display_zoom_level:true,
			event_modal:{href:'', type:'default'}
		},
		
		_create : function () {
		
			this._id = $(this.element).attr("id");
			/*
			Anatomy:
			*
			*  -container: main frame of entire timeline
			*  -centerline
			*  -truck: entire movable (left-right) container
			*  -ticks: includes "ruler" as well as events
			*  -handle: the grabbable part of the truck which 
			*           self-adjusts to center
			*  -slider-container: wrapper for zoom slider
			*  -slider: jQuery UI vertical slider
			*  -timeline-menu
			*
			*  -measure-span: utility div for measuring text lengths
			*
			*  -footer: (not shown) gets added dynamically unless
			*           options indicate otherwise
			*/
			// no need for template here as no data being passed
			var MAIN_TEMPLATE = "<div class='timeglider-container'>"
				+ "<div class='timeglider-loading'>loading</div>"
				+ "<div class='timeglider-centerline'></div>"
				+ "<div class='timeglider-date-display'></div>"
				+ "<div class='timeglider-truck' id='tg-truck'>"
				+ "<div class='timeglider-ticks'>"
				+ "<div class='timeglider-handle'></div>"
				+ "</div>"
				+ "</div>"
				+ "<div class='timeglider-slider-container'>"
				+ "<div class='timeglider-slider'></div>"
				+ "<div class='timeglider-pan-buttons'>"
				+ "<div class='timeglider-pan-left'></div><div class='timeglider-pan-right'></div>"
				+ "</div>"
				+ "</div>"
				+ "<div class='timeglider-footer'>"
				+ "<div class='timeglider-logo'></div>"                      
				+ "<div class='timeglider-footer-button timeglider-filter-bt'></div>"
				+ "<div class='timeglider-footer-button timeglider-settings-bt'></div>"
				+ "<div class='timeglider-footer-button timeglider-list-bt'></div>"
				+ "</div>"
				+ "<div class='timeglider-event-hover-info'></div>"
				+ "</div><span id='timeglider-measure-span'></span>";
			
			this.element.html(MAIN_TEMPLATE);
		
		}, // eof _create()
		
		/**
		* takes the created template and inserts functionality
		*  from Mediator and View constructors
		*
		*
		*/
		_init : function () {
		
			// validateOptions should come out as empty string
			var optionsCheck = timeglider.validateOptions(this.options);
			
			if (optionsCheck == "") {
			
				tg.TG_Date.setCulture(this.options.culture);
			
				MED = new tg.TG_Mediator(this.options);
				timelinePlayer = new tg.TG_PlayerView(this, MED);
			
				// after timelinePlayer is created this stuff can be done
				MED.setFocusDate(new TG_Date(this.options.initial_focus));
				
				
				MED.loadTimelineData(this.options.data_source);
			
			} else {
				alert("Rats. There's a problem with your widget settings:" + optionsCheck);
			}
		
		},
		
		
		/** 
		*********  PUBLIC METHODS ***************
		*
		*/
		
		
		/* 
		* goTo
		* sends timeline to a specific date and, optionally, zoom
		* @param d {String} ISO8601 date: 'YYYY-MM-DD HH:MM:SS'
		* @param z {Number} zoom level to change to; optional
		*/
		goTo : function (d, z) {
			MED.gotoDateZoom(d,z);
		},
		
		
		
		
		getMediator : function () {
			return MED;
		},
		
		
		/**
		* zoom
		* zooms the timeline in or out, adding an amount, often 1 or -1
		*
		* @param n {number|string}
		*          numerical: -1 (or less) for zooming in, 1 (or more) for zooming out
		*          string:    "in" is the same as -1, "out" the same as 1
		*/
		zoom : function (n) {
		
			switch(n) {
				case "in": n = -1; break;
				case "out": n = 1; break;
			}
			// non-valid zoom levels
			if (n > 99 || n < -99) { return false; }
			
			MED.zoom(n);
		},
		
		
		/**
		* zoom
		* zooms the timeline in or out, adding an amount, often 1 or -1
		*
		* @param n {number|string}
		*          numerical: -1 (or less) for zooming in, 1 (or more) for zooming out
		*          string:    "in" is the same as -1, "out" the same as 1
		*/
		load : function (src) {
			MED.loadTimelineData(src);
		},
		
		
		
		/**
		*  panButton
		*  sets a pan action on an element for mousedown and mouseup|mouseover
		*  
		*
		*/
		panButton : function (sel, vel) {
			var _vel = 0;
			switch(vel) {
				case "left": _vel = 30; break;
				case "right": _vel = -30; break;
				default: _vel = vel; break;
			}
			timelinePlayer.setPanButton(sel, _vel);
		},
		
		
		/**
		* destroy 
		* wipes out everything
		*/
		destroy : function () {
			$.Widget.prototype.destroy.apply(this, arguments);
			$(this.element).html("");
		}
	
	}); // end widget process

})(jQuery);
