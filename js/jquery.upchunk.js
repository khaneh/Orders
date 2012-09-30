(function() {
  ;

  var __bind = function(fn, me){ return function(){ return fn.apply(me, arguments); }; };

  (function($, window, document) {
    var Plugin, defaults, pluginName;
    pluginName = 'upchunk';
    defaults = {
      url: '',
      chunk: false,
      fallback_id: '',
      chunk_size: 1024,
      file_param: 'file',
      name_param: 'file_name',
      max_file_size: 0,
      queue_size: 2,
      processNextImmediately: false,
      data: {},
      drop: function() {},
      dragEnter: function() {},
      dragOver: function() {},
      dragLeave: function() {},
      docEnter: function() {},
      docOver: function() {},
      docLeave: function() {},
      beforeEach: function() {},
      afterAll: function() {},
      rename: function(file) {
        return file.name;
      },
      error: function(err) {
        return alert(err);
      },
      fileAdded: function() {},
      uploadStarted: function() {},
      uploadFinished: function() {},
      progressUpdated: function() {}
    };
    Plugin = (function() {

      function Plugin(element, options) {
        this.element = element;
        this.docLeave = __bind(this.docLeave, this);

        this.docOver = __bind(this.docOver, this);

        this.docEnter = __bind(this.docEnter, this);

        this.docDrop = __bind(this.docDrop, this);

        this.dragLeave = __bind(this.dragLeave, this);

        this.dragOver = __bind(this.dragOver, this);

        this.dragEnter = __bind(this.dragEnter, this);

        this.drop = __bind(this.drop, this);

        this.process = __bind(this.process, this);

        this.opts = $.extend({}, defaults, options);
        this._defaults = defaults;
        this._name = pluginName;
        this.errors = {
          notSupported: 'BrowserNotSupported',
          tooLarge: 'FileTooLarge',
          uploadHalted: 'UploadHalted'
        };
        this.hash = function(s) {
          var char, hash, i, len, test, _i;
          hash = 0;
          len = s.length;
          if (len === 0) {
            return hash;
          }
          for (i = _i = 0; 0 <= len ? _i <= len : _i >= len; i = 0 <= len ? ++_i : --_i) {
            char = s.charCodeAt(i);
            test = ((hash << 5) - hash) + char;
            if (!isNaN(test)) {
              hash = test & test;
            }
          }
          return Math.abs(hash);
        };
        this.init();
      }

      Plugin.prototype.init = function() {
        var _this = this;
        $(this.element).on('drop', this.drop).on('dragenter', this.dragEnter).on('dragover', this.dragOver).on('dragleave', this.dragLeave);
        $(document).on('drop', this.docDrop).on('dragenter', this.docEnter).on('dragover', this.docOver).on('dragleave', this.docLeave);
        return $('#' + this.opts.fallback_id).change(function(e) {
          var file, hash, i, _i, _j, _len, _ref, _ref1;
          _this.docLeave(e);
          _this.opts.drop(e);
          _this.files = e.target.files;
          if (!_this.files) {
            _this.opts.error(_this.errors.notSupported);
            false;
          }
          _this.processQ = [];
          _this.todoQ = [];
          _ref = _this.files;
          for (_i = 0, _len = _ref.length; _i < _len; _i++) {
            file = _ref[_i];
            _this.todoQ.push(file);
          }
          for (i = _j = 0, _ref1 = _this.opts.queue_size; 0 <= _ref1 ? _j <= _ref1 : _j >= _ref1; i = 0 <= _ref1 ? ++_j : --_j) {
            file = _this.todoQ.pop();
            if (file) {
              hash = (_this.hash(file.name) + file.size).toString();
              _this.opts.uploadStarted(file, hash);
              _this.processQ.push(file);
              _this.process(i);
            }
          }
          e.preventDefault();
          return false;
        });
      };

      Plugin.prototype.process = function(i) {
        var chunk_size, chunks, file, n, next_chunk, next_file, progress, send,
          _this = this;
        next_file = function() {
          var next, next_hash;
          next = _this.todoQ.pop();
          if (next) {
            next_hash = (_this.hash(next.name) + next.size).toString();
            _this.opts.uploadStarted(next, next_hash);
            _this.processQ.splice(i, 1, next);
            return _this.process(i);
          } else {
            return _this.processQ.splice(i, 1, false);
          }
        };
        progress = function(e) {
          var hash, old, pchunk, percentage;
          if (!(typeof old !== "undefined" && old !== null)) {
            old = 0;
          }
          if (typeof n !== "undefined" && n !== null) {
            pchunk = chunk_size * 100 / file.size;
            percentage = Math.floor(((e.loaded * 100) / file.size) + (n - 1) * pchunk);
          } else {
            percentage = Math.floor(e.loaded * 100 / e.total);
          }
          if (percentage > 100) {
            percentage = 100;
          }
          if (percentage > old) {
            old = percentage;
            hash = (_this.hash(file.name) + file.size).toString();
            _this.opts.progressUpdated(file, hash, percentage);
          }
          if (percentage === 100 && _this.opts.processNextImmediately) {
            return next_file();
          }
        };
        next_chunk = function() {
          var chunk, chunks, end, start;
          start = chunk_size * n;
          end = chunk_size * (n + 1);
          n += 1;
          if (file.slice) {
            chunk = file.slice(start, end);
          } else if (file.mozSlice) {
            chunk = file.mozSlice(start, end);
          } else if (file.webkitSlice) {
            chunk = file.webkitSlice(start, end);
          } else {
            chunk = file;
            chunks = 1;
          }
          return send(chunk, _this.opts.url);
        };
        send = function(chunk, url) {
          var fd, hash, name, value, xhr, _ref;
          hash = (_this.hash(file.name) + file.size).toString();
          if (_this.opts.beforeEach() === false) {
            _this.opts.error(_this.errors.uploadHalted);
            next_file();
            return false;
          }
          fd = new FormData;
          fd.append(_this.opts.file_param, chunk);
          fd.append(_this.opts.name_param, _this.opts.rename(file));
          fd.append('hash', hash);
          _ref = _this.opts.data;
          for (name in _ref) {
            value = _ref[name];
            if (typeof value === 'function') {
              fd.append(name, value());
            } else {
              fd.append(name, value);
            }
          }
          if (n === 1) {
            fd.append('first', true);
          } else {
            fd.append('first', false);
          }
          if (n === chunks) {
            fd.append('last', true);
          } else {
            fd.append('last', false);
          }
          xhr = new XMLHttpRequest;
          xhr.open('POST', url, true);
          xhr.upload.addEventListener('progress', progress, false);
          xhr.send(fd);
          return xhr.onload = function() {
            var f, fin, response;
            if ((typeof chunks !== "undefined" && chunks !== null) && n < chunks) {
              next_chunk();
            } else {
              try {
                response = $.parseJSON(xhr.responseText);
              } catch (_error) {}
              if (response != null) {
                _this.opts.uploadFinished(file, hash, response);
              } else {
                _this.opts.uploadFinished(file, hash);
              }
              if (!_this.opts.processNextImmediately) {
                next_file();
              }
            }
            fin = (function() {
              var _i, _len, _ref1, _results;
              _ref1 = this.processQ;
              _results = [];
              for (_i = 0, _len = _ref1.length; _i < _len; _i++) {
                f = _ref1[_i];
                if (f) {
                  _results.push(f);
                }
              }
              return _results;
            }).call(_this);
            if (fin.length === 0) {
              return _this.opts.afterAll();
            }
          };
        };
        file = this.processQ[i];
        if (this.opts.max_file_size > 0 && file.size > 1048576 * this.opts.max_file_size) {
          this.opts.error(this.errors.tooLarge);
          next_file();
          return false;
        }
        if (this.opts.chunk === false) {
          return send(file, this.opts.url);
        } else {
          chunk_size = 1024 * this.opts.chunk_size;
          chunks = Math.ceil(file.size / chunk_size);
          n = 0;
          return next_chunk();
        }
      };

      Plugin.prototype.drop = function(e) {
        var file, hash, i, _i, _j, _len, _ref, _ref1;
        this.docLeave(e);
        this.opts.drop(e);
        this.files = e.originalEvent.dataTransfer.files;
        if (!this.files) {
          this.opts.error(this.errors.notSupported);
          false;
        }
        this.processQ = [];
        this.todoQ = [];
        _ref = this.files;
        for (_i = 0, _len = _ref.length; _i < _len; _i++) {
          file = _ref[_i];
          this.todoQ.push(file);
        }
        for (i = _j = 0, _ref1 = this.opts.queue_size; 0 <= _ref1 ? _j <= _ref1 : _j >= _ref1; i = 0 <= _ref1 ? ++_j : --_j) {
          file = this.todoQ.pop();
          if (file) {
            hash = (this.hash(file.name) + file.size).toString();
            this.opts.uploadStarted(file, hash);
            this.processQ.push(file);
            this.process(i);
          }
        }
        e.preventDefault();
        return false;
      };

      Plugin.prototype.dragEnter = function(e) {
        clearTimeout(this.timer);
        e.preventDefault();
        return this.opts.dragEnter(e);
      };

      Plugin.prototype.dragOver = function(e) {
        clearTimeout(this.timer);
        e.preventDefault();
        this.opts.docOver(e);
        return this.opts.dragOver(e);
      };

      Plugin.prototype.dragLeave = function(e) {
        clearTimeout(this.timer);
        this.opts.dragLeave(e);
        return e.stopPropagation();
      };

      Plugin.prototype.docDrop = function(e) {
        e.preventDefault();
        this.opts.docLeave(e);
        return false;
      };

      Plugin.prototype.docEnter = function(e) {
        clearTimeout(this.timer);
        e.preventDefault();
        this.opts.docEnter(e);
        return false;
      };

      Plugin.prototype.docOver = function(e) {
        clearTimeout(this.timer);
        e.preventDefault();
        this.opts.docOver(e);
        return false;
      };

      Plugin.prototype.docLeave = function(e) {
        var _this = this;
        return this.timer = setTimeout((function() {
          return _this.opts.docLeave(e);
        }), 200);
      };

      return Plugin;

    })();
    return $.fn[pluginName] = function(options) {
      return this.each(function() {
        if (!$.data(this, "plugin_" + pluginName)) {
          return $.data(this, "plugin_" + pluginName, new Plugin(this, options));
        }
      });
    };
  })(jQuery, window, document);

}).call(this);

