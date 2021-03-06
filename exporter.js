Ext.define("Niks.Apps.TreeExporter", {
    //dateFormat : 'Y-m-d g:i',
    dateFormat : 'Y-m-d',

    config: {
    },

    constructor: function (config) {
        this.mergeConfig(config);
        this.callParent(arguments);
    },

    _downloadFiles: function( files ) {
        if ( files.length )
        {
            var data = files.pop();
            var format = files.pop();
            var file = files.pop();

            var href = "<a href='" + format + "," + encodeURIComponent(data) + "' download='" + file + "'></a>";

            var ml = Ext.DomHelper.insertAfter(window.document.getElementsByClassName('app')[0], href);
            ml.click();
//            ml.remove(); //Leaves them behind without this, but there is a timing issue: from click to remove.
            this._downloadFiles(files);
        }
    },

    exportCSV: function(tree) {
        var data = this._getCSV(tree);
        // fix: ' character was causing termination of csv file
        data = data.replace(/\'/g, " ");
        this._downloadFiles(
            [
                'export.csv', 'data:text/csv;charset=utf8', data
            ]
        );

    },

    _XMLIndent: function(index, tag, leaf, data) {
        var text = '\n';
        for (var i = 0; i < index; i++) { text += '\t'; }
        text += '<' + tag + '>';
        text += data;
        if (!leaf) { text += '\n'; for (i = 0; i < index; i++) {text += '\t';}}
        text += '</' + tag + '>';
        return text;

    },

    _escapeForCSV: function(string) {
        string = "" + string;
        if (string.match(/,/)) {
            if (!string.match(/"/)) {
                string = '"' + string + '"';
            } else {
                string = string.replace(/,/g, ''); // comma's and quotes-- sorry, just lose the commas
            }
        }
        return string;
    },

    _getFieldText: function(fieldData) {
        var text;
        if (fieldData === null || fieldData === undefined) {
            text = '';

        // we capture object types here
        } else if (fieldData._refObjectName && !fieldData.getMonth) {
            text = fieldData._refObjectName;

        // Date types here
        } else if (fieldData instanceof Date) {
            text = Ext.Date.format(fieldData, this.dateFormat);

        //The dependencies field is a synthetic one. We identity it by the field contents
        } else if (fieldData.Count !== undefined) {
            text = fieldData.Count.toString();
        }
        /*else if (!fieldData.match) { // not a string or object we recognize...bank it out
            text = '';
        } */ else {
            var delimiter = ",",
            rowDelimiter = "\r\n",
            re = new RegExp(delimiter + '|\"|\r|\n','g'),
            reHTML = new RegExp('<\/?[^>]+>', 'g'),
            reNbsp = new RegExp('&nbsp;','ig');

            text = fieldData;
            if (reHTML.test(text)){
                text = fieldData.replace('<br>',rowDelimiter);
                text = Ext.util.Format.htmlDecode(text);
                text = Ext.util.Format.stripTags(text);
            }
            if (reNbsp.test(text)){
                text = text.replace(reNbsp,' ');
            }

            if (re.test(text)){ //enclose in double quotes if we have the delimiters
                text = text.replace(/\"/g,'\"\"');
                text = Ext.String.format("\"{0}\"",text);
            }
        }

        return text;
    },

    _getFieldTextAndEscape: function(fieldData) {
        var string  = this._getFieldText(fieldData);

        return this._escapeForCSV(string);
    },

    // have to add the colIdx to the count of locked columns
    fixedColumnCount : function(columns) {
        var cols = _.filter(columns,function(c) { 
            return c!==undefined && c!==null && c.locked === true;
        });
        return cols.length;
    },

    traverseChildren: function(item, fields) {

        var me = this;
        var bigString = '';
        if (item.children) {
            _.each(item.children, function(child) { bigString += me.traverseChildren(child, fields);});
        }

        //If we come here, we are done going downwards
        return me.stringifyItem(item, fields) + bigString;
    },

    stringifyItem: function(item, fields) {
        var level = 1;
        var smallString = '';
        var parent = item.parent;
        var me = this;

        while (item.depth > level) {
            parent = parent.parent;
            smallString += ",";
            level +=1;
        }
        if (item.depth > 0) {
            smallString += item.data.record.get("FormattedID");
            level = parent.height - item.depth;
            while (level > 0) { 
                smallString += ",";
                level -=1;
            }
            smallString += ',' + this._getFieldTextAndEscape(item.data.record.data.Name);

            //Add more fields in here before attachments

            _.each(fields, function(field) {
                smallString += ',' + me._getFieldTextAndEscape(item.data.record.data[field]);    
            });

            if ( _.contains(this.fields, 'Attachments')){

                smallString += ',' + (item.data.Attachments?item.data.Attachments.Count:0);
                smallString += ',' + (item.data.Attachments?item.data.Attachments.Size:0);
                smallString += ',' + item.ChildAttachments.Size;
            }
        }
        return smallString + '\n';
    },

    _getCSV: function (tree) {
        var textOut = '';
        var fieldRow = '';
        var specialFields = ['Attachments', 'FormattedID', 'Name'];    //Dealt with individually in this code

        var printFields = _.filter(this.fields, function(field) {
            return !(_.contains(specialFields, field));
        });

        var that = this;
        if (!tree) { return; }  //If user clicks export before the tree is ready.....

        //First, let's generate the Column header row

        //The 'tree' item is the root item, but there is inconsistency in the types lower down Feature->Story->Defect or Feature->Story->Task
        var path = _.find( tree.descendants(), function(item) {
            return item.depth === tree.height;
        }).path(tree);
        path.pop(); //Get rid of R0
        var i = 0;
        for ( i = 0; i < tree.height; i++ ) {
            var pathItem = path.pop();
            console.log('Path level: ' + pathItem.height, 'Name: ' + pathItem.id, 'Record: ',pathItem.data.record);
            var title = pathItem.data.record.raw._type;
            if ( title.startsWith('PortfolioItem/')){
                title = title.replace('PortfolioItem/', '');
            }
            else { title = '';}
            fieldRow += title + ',';
        }

        fieldRow += 'Name';  //Mandatory in exporter
        _.each(printFields, function(field) {
            fieldRow += ',' + field;
        });
        
        if ( _.contains(this.fields, 'Attachments')){
            fieldRow += ",Attachments Count, Attachments Size, Attachments Total" ;
        }
        fieldRow += "\n";

        //Now we can start traversing the children
        textOut = that.traverseChildren(tree, printFields) + textOut;

        if (tree.ChildAttachments.Count > 0) {
            for ( i = 0; i < tree.height; i++){
                textOut += ',';
            }
            textOut += "Total Attachments Size";
            for ( i = 0; i < printFields.length; i++) {
                textOut+=',';
            }

            textOut += "," + tree.ChildAttachments.Count + ',' + tree.ChildAttachments.Size + '\n';
        }
        if (textOut.length > 0) {
            return fieldRow + textOut;
        }
        else {
            return null;
        }
    }
});