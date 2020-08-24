var inProgress = false;
function init() {
    error_message = "Internal error occured.";
    selected_node = null;
    icon = [];
    icon = top.icon;
    nodes = [];
    nodes_a = [];
    var category;
    var hasChildren;

    id = 0;
    var pathInXML = "";
    var ml = document.location.href;
    var root_name = null;
    mainPath = ml.substring(0, ml.lastIndexOf('treeview'));

    var xmlFile = "top.xml";

    var xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
    xml.validateOnParse = false;

    xml.async = false;
    try {
        xml.load(xmlFile);
        var oNode = xml.selectSingleNode("node").text;

        var t = "" + mainPath + oNode + "/path.xml";
        var xmlDoc = new ActiveXObject("MSXML2.DOMDocument.3.0");
        xmlDoc.async = false;
        xmlDoc.validateOnParse = false;
        var parent_name = "";
        var level;
        if (t != "") {
            xmlDoc.load(t);
            if (xmlDoc.parseError.errorcode == null) {
                var root = xmlDoc.documentElement;
                var oNodeList = root.childNodes;
                var child = oNodeList(0).getAttribute("name");
                level = oNodeList.length;
                root_name = child;
                root_id = oNode;
                if (oNodeList.length > 1)
                    var parent_name = oNodeList(1).getAttribute("name");
                for (var i = 0; i < oNodeList.length - 1; i++) {
                    pathInXML = oNodeList(i).getAttribute("name") + '\\' + pathInXML;
                }
                // document.writeln(pathInXML);
            }
        }

        xmlFile = "" + mainPath + oNode + "/stat.xml";
        xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
        xml.validateOnParse = false;

        xml.async = false;

        xml.load(xmlFile);
        if (xml.parseError.errorcode == null) {
            var statRoot = xml.documentElement;
            var statList = statRoot.childNodes;
            if (statList.length > 0)
                var warn = parseInt(statList(0).getAttribute("converted"));
            else warn = 100;
        }

        xmlFile = "" + mainPath + oNode + "/cat.xml";
        xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
        xml.validateOnParse = false;

        xml.async = false;

        xml.load(xmlFile);
        if (xml.parseError.errorcode == null) {
            var categoryRoot = xml.documentElement;
            var categoryList = categoryRoot.childNodes;
            category = categoryList(categoryList.length - 1).getAttribute("category");
        }

        xmlFile = "" + mainPath + oNode + "/children.xml";
        xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
        xml.validateOnParse = false;

        xml.async = false;

        xml.load(xmlFile);
        if (xml.parseError.errorcode == null) {
            var childrenRoot = xml.documentElement;
            var childrenList = childrenRoot.childNodes;
            if (childrenList.length > 0)
                hasChildren = 'Y';
        }

        xmlFile = "" + mainPath + oNode + "/node.xml";
        xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
        xml.validateOnParse = false;

        xml.async = false;

        xml.load(xmlFile);
        if (xml.parseError.errorcode == null) {
            var url = xml.selectSingleNode("node").text;

            var myNode = new Node(root_id, -1, root_name, url);

            if (warn != 100)
                myNode.warning = true;

            myNode.level = level;

            if ((myNode.level == 1) || (myNode.level == 3) || (myNode.level == 5) || (myNode.level == 7))
                category = "category";
            /*
            myNode.icon = (myNode.warning)?icon.folder_error : icon.folder;
            else
            switch (parent_name)
            {
            case "Procedures":
            myNode.icon = (myNode.warning)?icon.procedure_error:icon.procedure;       
            break;
            case "Schemas":
            myNode.icon = (myNode.warning)? icon.schema_error : icon.schema;
            break;
            case "Views":
            myNode.icon = (myNode.warning)? icon.view_error : icon.view;
            break;
            case "Databases":
            myNode.icon = (myNode.warning)? icon.database_error : icon.database;
            break;
            case "Indexes":
            myNode.icon = (myNode.warning)? icon.index_error : icon.index;
            break;
            case "Triggers":
            myNode.icon = (myNode.warning)? icon.trigger_error : icon.trigger;
            break;
            case "Tables":
            myNode.icon = (myNode.warning)? icon.table_error : icon.table;
            break;
            case "Queries":
            myNode.icon = (myNode.warning)? icon.query_error : icon.query;
            break;
            default:
            myNode.icon = (myNode.warning)?icon.folder_error : icon.folder;
            break;
            }            */
            myNode._id = id;
            nodes[id] = myNode;
            nodes_a[myNode.id] = myNode;
            level = 0;
            var myNode_str = createNode(myNode, id, warn, category, hasChildren);

            document.write(myNode_str);
        }
    }
    catch (e) {
        document.write(error_message);
        window.parent.nav.location = "about:blank";
        window.parent.node.location = "about:blank";
        window.parent.messages.location = "about:blank";

    }
    changeSelection(id);
    id++;
}

function getSelected() {
    return selected_node;
}

function getSelectedNodeId() {
    return nodes[selected_node].id;
}

function setIcon(node, category) {
    switch (category) {
        case "category":
            node.icon = (node.warning) ? icon.folder_error : icon.folder;
            break;
        case "packaged-procedure":
        case "procedure":
        case "object-member-procedure":
        case "object-static-procedure":
            node.icon = (node.warning) ? icon.procedure_error : icon.procedure;
            break;
        case "schema":
            node.icon = (node.warning) ? icon.schema_error : icon.schema;
            break;
        case "table":
            node.icon = (node.warning) ? icon.table_error : icon.table;
            break;
        case "view":
            node.icon = (node.warning) ? icon.view_error : icon.view;
            break;
        case "trigger":
            node.icon = (node.warning) ? icon.trigger_error : icon.trigger;
            break;
        case "index":
            node.icon = (node.warning) ? icon.index_error : icon.index;
            break;
        case "unique-constraint":
            node.icon = (node.warning) ? icon.unique_key_error : icon.unique_key;
            break;
        case "query":
            node.icon = (node.warning) ? icon.query_error : icon.query;
            break;
        case "mdb-file":
        case "database":
            node.icon = (node.warning) ? icon.database_error : icon.database;
            break;
        case "primary-key":
            node.icon = (node.warning) ? icon.primary_key_error : icon.primary_key;
            break;
        case "foreign-key":
            node.icon = (node.warning) ? icon.foreign_key_error : icon.foreign_key;
            break;
        case "packaged-function":
        case "function":
        case "object-member-function":
        case "object-static-function":
            node.icon = (node.warning) ? icon.function_error : icon.function_gif;
            break;
        case "package":
            node.icon = (node.warning) ? icon.package_error : icon.package_gif;
            break;
        case "synonym":
            node.icon = (node.warning) ? icon.synonym_error : icon.synonym_gif;
            break;
        case "sequence":
            node.icon = (node.warning) ? icon.sequence_error : icon.sequence_gif;
            break;
        case "private-packaged-subtype":
        case "private-packaged-type":
            node.icon = (node.warning) ? icon.private_type_error : icon.private_type_gif;
            break;
        case "packaged-type":
        case "packaged-subtype":
            node.icon = (node.warning) ? icon.type_error : icon.type_gif;
            break;
        case "private-packaged-procedure":
            node.icon = (node.warning) ? icon.private_procedure_error : icon.private_procedure_gif;
            break;
        case "private-packaged-function":
            node.icon = (node.warning) ? icon.private_function_error : icon.private_function_gif;
            break;
        case "statement":
            node.icon = (node.warning) ? icon.statement_error : icon.statement_gif;
            break;
        case "user-defined-type":
            node.icon = (node.warning) ? icon.datatype_error : icon.datatype_gif;
            break;
        case "object-type":
            node.icon = (node.warning) ? icon.object_type_error : icon.object_type_gif;
            break;
        case "collection-type":
            node.icon = (node.warning) ? icon.collection_type_error : icon.collection_type_gif;
            break;
        case "object-constructor-function":
            node.icon = (node.warning) ? icon.collection_type_error : icon.collection_type_gif;
            break;
    }
}

/*function setIcon(node,node_id)
{
var parent_id;
    
for (i=0;i<nodes.length;i++)
if (nodes[i].id == nodes[node_id].pid)
{
parent_id = i;
break;
}

if ((node.level == 1)||(node.level == 3)||(node.level == 5)||(node.level == 7))
node.icon = (node.warning)?icon.folder_error : icon.folder;
else
if (parent_id||parent_id == 0)
{
    
if (nodes[parent_id].name == "Procedures")
node.icon = (node.warning)? icon.procedure_error : icon.procedure;
else if ((nodes[parent_id].name == "Schemas"))
node.icon = (node.warning)? icon.schema_error : icon.schema;
else if ((nodes[parent_id].name == "Databases"))
node.icon = (node.warning)? icon.database_error : icon.database;
else if ((nodes[parent_id].name == "Tables"))
node.icon = (node.warning)? icon.table_error : icon.table;
else if ((nodes[parent_id].name == "Indexes"))
{
xmlFile = ""+mainPath +nodes[node_id].id+"/cat.xml";
xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
xml.validateOnParse = false;

xml.async = false;

xml.load(xmlFile);
if (xml.parseError.errorcode == null) 
{
var root = xml.documentElement;
var oNodeList = root.childNodes;
var category = oNodeList(0).getAttribute("category");
if (category == "unique-constraint")
node.icon = (node.warning)? icon.unique_key_error : icon.unique_key;
else
node.icon = (node.warning)? icon.index_error : icon.index;
}
}
else if ((nodes[parent_id].name == "Triggers"))
node.icon = (node.warning)? icon.trigger_error : icon.trigger;
else if ((nodes[parent_id].name == "Views"))
node.icon = (node.warning)? icon.view_error : icon.view;
else if ((nodes[parent_id].name == "Keys"))
{       
xmlFile = ""+mainPath +nodes[node_id].id+"/cat.xml";
xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
xml.validateOnParse = false;

xml.async = false;

xml.load(xmlFile);
if (xml.parseError.errorcode == null) 
{
var root = xml.documentElement;
var oNodeList = root.childNodes;
var category = oNodeList(0).getAttribute("category");
if (category == "primary-key")
node.icon = (node.warning) ? icon.primary_key_error : icon.primary_key;
else
node.icon = (node.warning)? icon.foreign_key_error : icon.foreign_key;
}

}
else if ((nodes[parent_id].name == "Constraints"))
{
xmlFile = ""+mainPath +nodes[node_id].id+"/cat.xml";
xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
xml.validateOnParse = false;

xml.async = false;

xml.load(xmlFile);
if (xml.parseError.errorcode == null) 
{
var root = xml.documentElement;
var oNodeList = root.childNodes;
var category = oNodeList(0).getAttribute("category");
if (category == "primary-key")
node.icon = (node.warning) ? icon.primary_key_error : icon.primary_key;
else
if (category == "unique-constraint")
node.icon = (node.warning)? icon.unique_key_error : icon.unique_key;
else
if (category == "check-constraint")
node.icon = (node.warning)? icon.check_constraint_error : icon.check_constraint;
else
node.icon = (node.warning)? icon.foreign_key_error : icon.foreign_key;
}
            
}
else if ((nodes[parent_id].name == "Columns"))
node.icon = (node.warning)?icon.column_error : icon.column;
else if ((nodes[parent_id].name == "Queries"))
node.icon = (node.warning)? icon.query_error : icon.query;
else if (node.warning)
node.icon = icon.folder_error;
else
node.icon = icon.folder;
}
}*/


function addNodes(parent_id) {
    var eChildren = document.getElementById("ch_id" + parent_id);
    if (eChildren == null) {
        var str = "<DIV id=ch_id" + parent_id + ">";
        var parent_node = nodes[parent_id];
        var t = "" + mainPath + parent_node.id + "/children.xml";
        var xmlDoc = new ActiveXObject("MSXML2.DOMDocument.3.0");
        xmlDoc.async = false;
        xmlDoc.validateOnParse = false;
        if (t != "") {
            xmlDoc.load(t);
            var root = xmlDoc.documentElement;
            var oNodeList = root.childNodes;
            var warn;
            var category;
            var hasChildren;
            var level = parent_node.level + 1;
            for (var i = 0; i < oNodeList.length; i++) {

                child = oNodeList(i).getAttribute("id");
                child_name = oNodeList(i).getAttribute("name");
                warn = parseInt(oNodeList(i).getAttribute("converted"));
                category = oNodeList(i).getAttribute("category");
                hasChildren = oNodeList(i).getAttribute("has-children");
                xmlFile = "" + mainPath + child + "/node.xml";
                xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
                xml.validateOnParse = false;
                xml.async = false;
                try {
                    xml.load(xmlFile);
                    if (xml.parseError.errorcode == null) {
                        var url = xml.selectSingleNode("node").text;
                        var new_node = new Node(child, parent_node.id, child_name, url);
                        if (warn != 100)
                            new_node.warning = true;
                        new_node.level = level;
                        new_node._id = id;
                        nodes[id] = new_node;
                        nodes_a[new_node.id] = new_node;
                        var new_node_str = createNode(new_node, id, warn, category, hasChildren);
                        id++;
                        str += new_node_str;
                    }
                }
                catch (e) {
                    var url = "about:blank";
                    var new_node = new Node(child, parent_node.id, child_name, url);
                    new_node.warning = true;
                    new_node.level = level;
                    new_node._id = id;
                    nodes[id] = new_node;
                    nodes_a[new_node.id] = new_node;

                    var new_node_str = createNodeWithoutChildren(new_node, id);
                    id++;
                    str += new_node_str;
                }


            }
            str += "</DIV>";
            var oDiv = document.createElement("DIV");
            oDiv.innerHTML = str;
            oDiv.style.marginLeft = 14;
            var parDiv = document.getElementById("node" + parent_id);
            parDiv.appendChild(oDiv);
        }
    }
    else
        eChildren.style.display = 'block';
}

function al() {
    inProgress = false;
    alert("SSMA: Object is not accessible.");
}

function createNodeWithoutChildren(node, nodeId) {
    var str = '<div class="dTreeNode" id="node' + nodeId + '">';

    node._hc = false;

    setIcon(node, nodeId);

    str += '<img id="j' + nodeId + '" src="';
    str += icon.empty;
    str += '" alt="" /></a>';
    if ((node.name == "Indexes") || (node.name == "Keys"))
        node.icon = (node.warning) ? icon.folder_error : icon.folder;


    if (root_id == node.id)
        switch (node.name) {
        case "Databases":
            node.icon = icon.source;
            break;
    }

    str += '<img id="i' + nodeId + '" src="' + ((node._io) ? node.iconOpen : node.icon) + '" alt="" />';



    if (node.url) {
        str += '<a id="s' + nodeId + '" class="node"' + '" href="javascript: al();"';
        str += '>';

    }

    str += node.name;


    if (node.url || (!node.url && node._hc))
        str += "</a>";

    str += '</div>';

    return str;
}


function removeNodes(id) {
    var eDiv = document.getElementById("ch_id" + id);
    eDiv.style.display = 'none';
}



function findParentById(id) {
    return nodes_a[nodes[id].pid];
}


function findParent(node) {
    return nodes_a[node.pid];
}

function getNodeId(node) {
    return nodes_a[node.id]._id;
}


//open node in the tree by node id
function NavigateTo(id) {
    try {
        //    var node_id=-1;
        var node_id = "00000000-0000-0000-0000-000000000000";
        //try to find node in the tree

        if (nodes_a[id] != null)
            node_id = nodes_a[id]._id;

        //if it exist open it and all it's parent nodes
        if (node_id != "00000000-0000-0000-0000-000000000000") {
            if (node_id != 0) {
                var par = findParentById(node_id);

                var nodesArray = [];
                var array_length = 1;
                nodesArray[0] = par;

                //add parent nodes to array
                while (par != nodes[0]) {
                    par = nodes_a[par.pid];
                    nodesArray[array_length] = par;
                    array_length++;
                }

                //do o function for all items from array
                for (var i = array_length - 1; i >= 0; i--)
                    if (nodesArray[i]._io == false)
                        o(getNodeId(nodesArray[i]));
            }
            inProgress = false;
            s(node_id);
            var eElem = document.getElementById('node' + node_id);
            eElem.scrollIntoView(true);
        }

        //if not, open path.xml and iterate throgh parents of current child
        else {
            var xml = new ActiveXObject("MSXML2.DOMDocument.3.0");
            xml.validateOnParse = false;
            xml.async = false;

            var path = mainPath + id + "/path.xml"
            xml.load(path);
            if (xml.parseError.errorcode == null) {
                var root = xml.documentElement;
                var oNodeList = root.childNodes;
                var nodesArray = [];
                var array_length = 0;

                for (var i = 1; i < oNodeList.length; i++) {
                    //	     parent_node_id = -1;
                    parent_node_id = "00000000-0000-0000-0000-000000000000";
                    var child = oNodeList[i].getAttribute("id");
                    if (nodes_a[child])
                        parent_node_id = nodes_a[child].id;
                    nodesArray[array_length] = child;
                    array_length++;
                    if (parent_node_id != "00000000-0000-0000-0000-000000000000")
                        break;
                }
                // when parent in the tree is found open all parent's children until we reach necessary node
                for (var i = array_length - 1; i >= 0; i--) {
                    o(nodes_a[nodesArray[i]]._id);
                }
                var open_id = nodes_a[id]._id;
                inProgress = false;
                s(open_id);
                var eElem = document.getElementById('node' + open_id);
                eElem.scrollIntoView(true);
            }
        }
    }
    catch (e) {
        inProgress = false;
        alert("SSMA: Object is not accessible.");
    }
}

function o(id) {
    document.body.style.cursor = "progress";
    var elemDiv = document.getElementById("exp" + id);
    //var node_text = elemDiv.innerText;
    //elemDiv.innerText = node_text + " (loading...) ";
    elemDiv.style.display = "block";
    try {
        var cn = nodes[id];
        if (cn._hc == true) {

            nodeStatus(!cn._io, id);

            cn._io = !cn._io;

            if (cn._io) {
                if (cn._hc) {
                    var eImg = document.getElementById("j" + id);
                    eImg.src = icon.nlMinus;
                }
                addNodes(id);

            }
            else {
                if (cn._hc) {
                    var eImg = document.getElementById("j" + id);
                    eImg.src = icon.nlPlus;
                }
                removeNodes(id);
            }
        }
    }
    catch (e) {
        document.write(error_message);
        window.parent.nav.location = "about:blank";
        window.parent.node.location = "about:blank";
        window.parent.messages.location = "about:blank";

    }
    elemDiv.style.display = "none";
    document.body.style.cursor = "auto";
}

function s(node_id) {
    if (inProgress)
        return;
    inProgress = true;
    window.parent.node.location = nodes[node_id].url;
    top.getNode(nodes[node_id].id);
    changeSelection(node_id);
}

function changeSelection(node_id) {
    if (selected_node != node_id) {
        if (selected_node != null) {
            var eSel = document.getElementById('s' + selected_node);
            eSel.className = 'node';
        }

        selected_node = node_id;
        var eNode = document.getElementById('s' + node_id);
        eNode.className = 'nodeSel';
    }
}

function nodeStatus(status, id) {

    eDiv = document.getElementById('d' + id);

    eDiv.style.display = (status) ? 'block' : 'none';

}


function checkForChildren(node) {
    var t = "" + mainPath + node.id + "/children.xml";
    var xmlDoc = new ActiveXObject("MSXML2.DOMDocument.3.0");
    xmlDoc.async = false;
    xmlDoc.validateOnParse = false;
    if (t != "") {
        xmlDoc.load(t);
        if (xmlDoc.parseError.errorcode == null) {
            var root = xmlDoc.documentElement;
            var oNodeList = root.childNodes;
            if (oNodeList.length > 0)
                return true;
        }
    }
    return false;
}

function createNode(node, nodeId, warn, category, hasChildren) {
    var str = '<div class="dTreeNode" style="width:auto;" id="node' + nodeId + '">';

    /*    if (warn<=100)
    node._hc = checkForChildren(node);
    else 
    node._hc = false;*/
    if (hasChildren == 'Y')
        node._hc = true;
    else
        node._hc = false;

    setIcon(node, category);

    if (node._hc) {
        str += '<a href="javascript: ' + 'o(' + nodeId + ');"><img id="j' + nodeId + '" src="';
        str += (node._io) ? icon.nlMinus : icon.nlPlus;
        str += '" alt="" /></a>';
    }
    else {
        str += '<img id="j' + nodeId + '" src="';
        str += icon.empty;
        str += '" alt="" />';
        /*if ((node.name == "Indexes")||(node.name == "Keys"))
        node.icon = (node.warning)?icon.folder_error : icon.folder;*/
    }


    if (root_id == node.id)
        switch (node.name) {
        case "Databases":
            node.icon = icon.source;
            break;
    }


    str += '<img id="i' + nodeId + '" src="' + node.icon + '" alt="" >';


    //	if (node.url) 
    {
        str += '<a style="width:auto;"  id="s' + nodeId + '" class="node"' + '" href="javascript: s(' + nodeId + ');">';

    }


    str += node.name;

    //	if (node.url)
    str += "</a>";
    str += "<div style='position:absolute;display:none;float:none' id='exp" + nodeId + "'>(expanding...)</div>";
    str += '</div>';

    if (node._hc) {
        str += '<div id="d' + nodeId + '" class="clip" style="display:' + ((root_id == node.pid || node._io) ? 'block' : 'none') + ';">';
        str += '</div>';
    }
    return str;
}


// Node object

function Node(id, pid, name, url, icon, iconOpen, open) {

    this.id = id;

    this.pid = pid;

    this.name = name;

    this.url = url;

    this.icon = icon;

    this.warning = false;

    this._io = open || false;

    this._hc = true;

    this._id = 0;

    this.level = 0;
};