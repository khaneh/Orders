function daydiff(first, second) {
    return (second-first)/(1000*60*60*24)
}
function getNum(n){
	var out = 0;
	if (typeof(n)=="number")
		out = parseInt(n);
	else if (typeof(n)=="string")
		if (n=="")
			out=0;
		else
			out = parseInt(n.replace(/,/gi,''));
	if (isNaN(out))
		out=0;
	return out;
}
function echoNum(str){
	var regex = /(-?[0-9]+)([0-9]{3})/;
	str = Math.floor(str);
    str += '';
    while (regex.test(str)) {
        str = str.replace(regex, '$1,$2');
    }
    //str += ' kr';
    return str;
}

$(document).ready(function(){
	$("input.date").blur(function(){
		var obj=$(this);
		obj.attr("title","");
		if (obj.val()=="") {
			obj.attr("title","·ÿ›«  «—ÌŒ —« Ê«—œ ﬂ‰Ìœ");
			obj.focus();
		}
		else if (obj.val()=="//") {
			var today = new Date();
			obj.val($.format.date(today,"yyyy/MM/dd"));
		} else {
			var rege=/^(13)?[7-9][0-9]\/[0-1]?[0-9]\/[0-3]?[0-9]$/;
			if( rege.test(obj.val()) ) {
				var SP = obj.val().split("/");
				if (SP[0].length == 2) SP[0] = "13" + SP[0] ;
				if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
				if (SP[2].length == 1) SP[2] = "0"  + SP[2] ;
				obj.val(SP.join("/"));	
			}
			if(!rege.test(obj.val())||( SP[0]<'1376' || SP[1]>'12' || SP[2]>'31' )) {
				obj.attr("title","›—„   «—ÌŒ »«Ìœ YYYY/MM/DD »«‘œ.");
				obj.focus();
			};
		}
	});
	$("input.time").blur(function(){
		var obj=$(this);
		obj.attr("title","");
		if (obj.val()=="") {
			obj.attr("title","·ÿ›« ”«⁄  —« Ê«—œ ﬂ‰Ìœ.");
			obj.focus();
		} else if (obj.val()==":"){
			var now = new Date();
			obj.val(now.getHours()+':'+now.getMinutes());
			var rege = /^[0-2]?[0-9]:[0-5]?[0-9]$/;
			if( rege.test(obj.val()) ) {
				var SP = obj.val().split(":");
				if (SP[0].length == 1) SP[0] = "0" + SP[0] ;
				if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
				obj.val(SP.join(":"));
			}
		} else {
			var rege = /^[0-2]?[0-9]:[0-5]?[0-9]$/;
			if( rege.test(obj.val()) ) {
				var SP = obj.val().split(":");
				if (SP[0].length == 1) SP[0] = "0" + SP[0] ;
				if (SP[1].length == 1) SP[1] = "0"  + SP[1] ;
				obj.val(SP.join(":"));
			}
			if(!rege.test(obj.val())||( SP[0]>23 || SP[1]>59)) {
				obj.attr("title","›—„  ”«⁄  »«Ìœ HH:MM »«‘œ.");
				obj.focus();
			};
		}
	
	});
	$("input.date").keypress(function(event){
/* 		console.log(event.keyCode); */
		if ( event.which == 47 || event.which == 8 || event.keyCode == 9 || event.which == 27 || event.which == 13 || 
             // Allow: Ctrl+A
            (event.which == 65 && event.ctrlKey === true) ||
            (event.keyCode == 37 || event.keyCode == 39)) {
                 // let it happen, don't do anything
                 return;
        }
        else {
            // Ensure that it is a number and stop the keypress
            if (event.shiftKey || (event.which < 48 || event.which > 57)) {
            	if (event.which>1775 && event.which<1786) 
		    		$(this).val($(this).val()+String.fromCharCode(event.which - 1728));
                event.preventDefault(); 
            }   
        }
	});
	$("input.time").keypress(function(event){
		if ( event.which == 58 || event.which == 8 || event.which == 9 || event.which == 27 || event.which == 13 || 
             // Allow: Ctrl+A
            (event.which == 65 && event.ctrlKey === true) ||
            (event.keyCode == 37 || event.keyCode == 39)) {
                 // let it happen, don't do anything
                 return;
        }
        else {
            // Ensure that it is a number and stop the keypress
            if (event.shiftKey || (event.which < 48 || event.which > 57)) {
            	if (event.which>1775 && event.which<1786) 
		    		$(this).val($(this).val()+String.fromCharCode(event.which - 1728));
                event.preventDefault(); 
            }   
        }
	});
	$("input:not(.num,.date)").keypress(function(event){
		if (event.which==1740){
			// change persian YEH to arabic YEH
			$(this).val($(this).val()+String.fromCharCode(1610));
			event.preventDefault(); 
		}
		if (event.which==1705){
			// change persian KE to arabic KE
			$(this).val($(this).val()+String.fromCharCode(1603));
			event.preventDefault(); 
		}
	});
    $(".num").keypress(function(event){
/*     	console.log(event.keyCode); */
    	// Allow: backspace, delete, tab, escape, and enter
        if ( event.which == 46 || event.which == 8 || event.which == 9 || event.which == 27 || event.which == 13 || 
             // Allow: Ctrl+A
            (event.which == 65 && event.ctrlKey === true) ||
            (event.keyCode == 37 || event.keyCode == 39)) {
                 // let it happen, don't do anything
                 return;
        }
        else {
            // Ensure that it is a number and stop the keypress
            if (event.shiftKey || (event.which < 48 || event.which > 57)) {
            	if (event.which>1775 && event.which<1786) 
		    		$(this).val($(this).val()+String.fromCharCode(event.which - 1728));
                event.preventDefault(); 
            }   
        }
    	
    });
    $("input.num").blur(function(){
	    $(this).val(echoNum(getNum($(this).val())));
    });
});
/* ===========================================================
 * bootstrap-tooltip.js v2.1.1
 * http://twitter.github.com/bootstrap/javascript.html#tooltips
 * Inspired by the original jQuery.tipsy by Jason Frame
 * ===========================================================
 * Copyright 2012 Twitter, Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * ========================================================== */


!function ($) {

  "use strict"; // jshint ;_;


 /* TOOLTIP PUBLIC CLASS DEFINITION
  * =============================== */

  var Tooltip = function (element, options) {
    this.init('tooltip', element, options)
  }

  Tooltip.prototype = {

    constructor: Tooltip

  , init: function (type, element, options) {
      var eventIn
        , eventOut

      this.type = type
      this.$element = $(element)
      this.options = this.getOptions(options)
      this.enabled = true

      if (this.options.trigger == 'click') {
        this.$element.on('click.' + this.type, this.options.selector, $.proxy(this.toggle, this))
      } else if (this.options.trigger != 'manual') {
        eventIn = this.options.trigger == 'hover' ? 'mouseenter' : 'focus'
        eventOut = this.options.trigger == 'hover' ? 'mouseleave' : 'blur'
        this.$element.on(eventIn + '.' + this.type, this.options.selector, $.proxy(this.enter, this))
        this.$element.on(eventOut + '.' + this.type, this.options.selector, $.proxy(this.leave, this))
      }

      this.options.selector ?
        (this._options = $.extend({}, this.options, { trigger: 'manual', selector: '' })) :
        this.fixTitle()
    }

  , getOptions: function (options) {
      options = $.extend({}, $.fn[this.type].defaults, options, this.$element.data())

      if (options.delay && typeof options.delay == 'number') {
        options.delay = {
          show: options.delay
        , hide: options.delay
        }
      }

      return options
    }

  , enter: function (e) {
      var self = $(e.currentTarget)[this.type](this._options).data(this.type)

      if (!self.options.delay || !self.options.delay.show) return self.show()

      clearTimeout(this.timeout)
      self.hoverState = 'in'
      this.timeout = setTimeout(function() {
        if (self.hoverState == 'in') self.show()
      }, self.options.delay.show)
    }

  , leave: function (e) {
      var self = $(e.currentTarget)[this.type](this._options).data(this.type)

      if (this.timeout) clearTimeout(this.timeout)
      if (!self.options.delay || !self.options.delay.hide) return self.hide()

      self.hoverState = 'out'
      this.timeout = setTimeout(function() {
        if (self.hoverState == 'out') self.hide()
      }, self.options.delay.hide)
    }

  , show: function () {
      var $tip
        , inside
        , pos
        , actualWidth
        , actualHeight
        , placement
        , tp

      if (this.hasContent() && this.enabled) {
        $tip = this.tip()
        this.setContent()

        if (this.options.animation) {
          $tip.addClass('fade')
        }

        placement = typeof this.options.placement == 'function' ?
          this.options.placement.call(this, $tip[0], this.$element[0]) :
          this.options.placement

        inside = /in/.test(placement)

        $tip
          .remove()
          .css({ top: 0, left: 0, display: 'block' })
          .appendTo(inside ? this.$element : document.body)

        pos = this.getPosition(inside)

        actualWidth = $tip[0].offsetWidth
        actualHeight = $tip[0].offsetHeight

        switch (inside ? placement.split(' ')[1] : placement) {
          case 'bottom':
            tp = {top: pos.top + pos.height, left: pos.left + pos.width / 2 - actualWidth / 2}
            break
          case 'top':
            tp = {top: pos.top - actualHeight, left: pos.left + pos.width / 2 - actualWidth / 2}
            break
          case 'left':
            tp = {top: pos.top + pos.height / 2 - actualHeight / 2, left: pos.left - actualWidth}
            break
          case 'right':
            tp = {top: pos.top + pos.height / 2 - actualHeight / 2, left: pos.left + pos.width}
            break
        }

        $tip
          .css(tp)
          .addClass(placement)
          .addClass('in')
      }
    }

  , setContent: function () {
      var $tip = this.tip()
        , title = this.getTitle()

      $tip.find('.tooltip-inner')[this.options.html ? 'html' : 'text'](title)
      $tip.removeClass('fade in top bottom left right')
    }

  , hide: function () {
      var that = this
        , $tip = this.tip()

      $tip.removeClass('in')

      function removeWithAnimation() {
        var timeout = setTimeout(function () {
          $tip.off($.support.transition.end).remove()
        }, 500)

        $tip.one($.support.transition.end, function () {
          clearTimeout(timeout)
          $tip.remove()
        })
      }

      $.support.transition && this.$tip.hasClass('fade') ?
        removeWithAnimation() :
        $tip.remove()

      return this
    }

  , fixTitle: function () {
      var $e = this.$element
      if ($e.attr('title') || typeof($e.attr('data-original-title')) != 'string') {
        $e.attr('data-original-title', $e.attr('title') || '').removeAttr('title')
      }
    }

  , hasContent: function () {
      return this.getTitle()
    }

  , getPosition: function (inside) {
      return $.extend({}, (inside ? {top: 0, left: 0} : this.$element.offset()), {
        width: this.$element[0].offsetWidth
      , height: this.$element[0].offsetHeight
      })
    }

  , getTitle: function () {
      var title
        , $e = this.$element
        , o = this.options

      title = $e.attr('data-original-title')
        || (typeof o.title == 'function' ? o.title.call($e[0]) :  o.title)

      return title
    }

  , tip: function () {
      return this.$tip = this.$tip || $(this.options.template)
    }

  , validate: function () {
      if (!this.$element[0].parentNode) {
        this.hide()
        this.$element = null
        this.options = null
      }
    }

  , enable: function () {
      this.enabled = true
    }

  , disable: function () {
      this.enabled = false
    }

  , toggleEnabled: function () {
      this.enabled = !this.enabled
    }

  , toggle: function () {
      this[this.tip().hasClass('in') ? 'hide' : 'show']()
    }

  , destroy: function () {
      this.hide().$element.off('.' + this.type).removeData(this.type)
    }

  }


 /* TOOLTIP PLUGIN DEFINITION
  * ========================= */

  $.fn.tooltip = function ( option ) {
    return this.each(function () {
      var $this = $(this)
        , data = $this.data('tooltip')
        , options = typeof option == 'object' && option
      if (!data) $this.data('tooltip', (data = new Tooltip(this, options)))
      if (typeof option == 'string') data[option]()
    })
  }

  $.fn.tooltip.Constructor = Tooltip

  $.fn.tooltip.defaults = {
    animation: true
  , placement: 'top'
  , selector: false
  , template: '<div class="tooltip"><div class="tooltip-arrow"></div><div class="tooltip-inner"></div></div>'
  , trigger: 'hover'
  , title: ''
  , delay: 0
  , html: true
  }

}(window.jQuery);
/* ============================================================
 * bootstrap-dropdown.js v2.1.1
 * http://twitter.github.com/bootstrap/javascript.html#dropdowns
 * ============================================================
 * Copyright 2012 Twitter, Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * ============================================================ */


!function ($) {

  "use strict"; // jshint ;_;


 /* DROPDOWN CLASS DEFINITION
  * ========================= */

  var toggle = '[data-toggle=dropdown]'
    , Dropdown = function (element) {
        var $el = $(element).on('click.dropdown.data-api', this.toggle)
        $('html').on('click.dropdown.data-api', function () {
          $el.parent().removeClass('open')
        })
      }

  Dropdown.prototype = {

    constructor: Dropdown

  , toggle: function (e) {
      var $this = $(this)
        , $parent
        , isActive

      if ($this.is('.disabled, :disabled')) return

      $parent = getParent($this)

      isActive = $parent.hasClass('open')

      clearMenus()

      if (!isActive) {
        $parent.toggleClass('open')
        $this.focus()
      }

      return false
    }

  , keydown: function (e) {
      var $this
        , $items
        , $active
        , $parent
        , isActive
        , index

      if (!/(38|40|27)/.test(e.keyCode)) return

      $this = $(this)

      e.preventDefault()
      e.stopPropagation()

      if ($this.is('.disabled, :disabled')) return

      $parent = getParent($this)

      isActive = $parent.hasClass('open')

      if (!isActive || (isActive && e.keyCode == 27)) return $this.click()

      $items = $('[role=menu] li:not(.divider) a', $parent)

      if (!$items.length) return

      index = $items.index($items.filter(':focus'))

      if (e.keyCode == 38 && index > 0) index--                                        // up
      if (e.keyCode == 40 && index < $items.length - 1) index++                        // down
      if (!~index) index = 0

      $items
        .eq(index)
        .focus()
    }

  }

  function clearMenus() {
    getParent($(toggle))
      .removeClass('open')
  }

  function getParent($this) {
    var selector = $this.attr('data-target')
      , $parent

    if (!selector) {
      selector = $this.attr('href')
      selector = selector && /#/.test(selector) && selector.replace(/.*(?=#[^\s]*$)/, '') //strip for ie7
    }

    $parent = $(selector)
    $parent.length || ($parent = $this.parent())

    return $parent
  }


  /* DROPDOWN PLUGIN DEFINITION
   * ========================== */

  $.fn.dropdown = function (option) {
    return this.each(function () {
      var $this = $(this)
        , data = $this.data('dropdown')
      if (!data) $this.data('dropdown', (data = new Dropdown(this)))
      if (typeof option == 'string') data[option].call($this)
    })
  }

  $.fn.dropdown.Constructor = Dropdown


  /* APPLY TO STANDARD DROPDOWN ELEMENTS
   * =================================== */

  $(function () {
    $('html')
      .on('click.dropdown.data-api touchstart.dropdown.data-api', clearMenus)
    $('body')
      .on('click.dropdown touchstart.dropdown.data-api', '.dropdown form', function (e) { e.stopPropagation() })
      .on('click.dropdown.data-api touchstart.dropdown.data-api'  , toggle, Dropdown.prototype.toggle)
      .on('keydown.dropdown.data-api touchstart.dropdown.data-api', toggle + ', [role=menu]' , Dropdown.prototype.keydown)
  })

}(window.jQuery);

/*
function setHeader(xhr) {

  xhr.setRequestHeader('Authorization', token);
}
function dial(num,exten){
	$.ajaxSetup({
		cache: false
	});
	$.ajax({
		type:"GET",
		url:"http://192.168.10.10:8088/pdhco/mxml",
		data:{action:"login",username:"mypdhco",secret:"66042700"},
		crossDomain: true,
		//headers:{'Access-Control-Allow-Origin':'*'},
		beforesend: {xhr.setRequestHeader('Access-Control-Allow-Origin', '*');},
		dataType:"XML"
	}).done(function (data){
		if (data.find("generic").attr("response")=="Success"){
			$.ajax({
				type:"GET",
				url:"http://192.168.10.10:8088/pdhco/mxml",
				data:{action:"Originate",Channel:"DAHDI/R1/" + num,Exten:exten,Context:"pdhco-cos4","Priority":"1",Async:"true"},
				headers:{'Access-Control-Allow-Origin':'*'},
				dataType:"XML"
			}).done(function (data){
				$.ajax({
					type:"GET",
					url:"http://192.168.10.10:8088/pdhco/mxml",
					data:{action:"logoff"},
					headers:{'Access-Control-Allow-Origin':'*'},
					dataType:"XML"
				}).done(alert("œ«Œ·Ì "+ exten + " ‘„«—Â " + num + "‘„«—Â êÌ—Ì „Ìùﬂ‰œ."));
			});
		}
	});
}
*/