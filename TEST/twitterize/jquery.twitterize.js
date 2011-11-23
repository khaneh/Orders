jQuery.fn.twitterize = function(options){
	var defaultSettings = {
		target: 'repsismx',
		cssClass: 'twitterize',
		statusNo: 5,
		divClass: 'userTimeline',
		innerDivClass: 'userTweet',
		linkClass: 'linkClass',
		showLink: true
	};
	var settings = jQuery.extend(defaultSettings, options);
	var timelineUrl = 'http://twitter.com/statuses/user_timeline/'+settings.target + '.rss'	;
	var tweets = new Array();
	var timestamp = new Date().getTime();
	$.jGFeed(timelineUrl,
	function(feeds){
		if(!feeds){
			return false;
		}
		for(var i = 0; i <= settings.statusNo; i++){
			var entry = feeds.entries[i];
			if(entry){
				tweets.push('<div class="' + settings.innerDivClass + '">' + entry.title + "</div>");
			}
		}
		printTweets();
	});
	function printTweets(){
	$('<div id="twitterize' + timestamp + '" class="userTimeline"></div>').appendTo($("body"));
	$('#twitterize' + timestamp).hide();
		for(var i = 0; i < tweets.length; i++){
			$('#twitterize' + timestamp).append(tweets[i]);
		}
		
	}
	return this.each(function(){
		$(this).addClass(settings.cssClass);
		$(this).hover(function(event){
		/** Mouse in */
			$('#twitterize' + timestamp).css('left', event.pageX);
			$('#twitterize' + timestamp).css('top', event.pageY);
			$('#twitterize' + timestamp).show();
		},
		function(event){
		/** Mouse out */
			$('#twitterize' + timestamp).hide();
		});
		if(settings.showLink){
			$(this).html('<a class="' + settings.linkClass + '" href="http://twitter.com/' + settings.target + '">' + 
							$(this).html());
			$(this).html($(this).html() + '</a>');
		}
	});
}