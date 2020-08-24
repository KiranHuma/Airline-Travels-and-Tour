function getIcon(objectType,messageType)
{

    switch (objectType)
    {
        case "table":
        {
            switch (messageType)
            {
                case "error":
                    return icon.table_error;
                case "warning":
                    return icon.table_warning;
                case "info":
                    return icon.table_info;
                default:
                    return icon.table;
            }
        }
        case "view":
        {
            switch (messageType)
            {
                case "error":
                    return icon.view_error;
                case "warning":
                    return icon.view_warning;
                case "info":
                    return icon.view_info;
                default:
                    return icon.view;
            }
        }
        
        case "index":
        {
            switch (messageType)
            {
                case "error":
                    return icon.index_error;
                case "warning":
                    return icon.index_warning;
                case "info":
                    return icon.index_info;
                default:
                    return icon.index;
            }
        }

        case "database":
        {
            switch (messageType)
            {
                case "error":
                    return icon.database_error;
                case "warning":
                    return icon.database_warning;
                case "info":
                    return icon.database_info;
                default:
                    return icon.database;
            }
        }
        
        case "query":
        {
            switch (messageType)
            {
                case "error":
                    return icon.query_error;
                case "warning":
                    return icon.query_warning;
                case "info":
                    return icon.query_info;
                default:
	            return icon.query;
            }
        }

    }
}    


icon = {

        primary_key         : 'img/Primary_Key.gif',
        
        primary_key_error   : 'img/primary_key_error.gif',
        
        foreign_key         : 'img/foreign_key.gif',
        
        foreign_key_error   : 'img/foreign_key_error.gif',
        
        unique_key          : 'img/unique_key.gif',
        
        unique_key_error    : 'img/unique_key_error.gif',
        
        column              : 'img/column.gif',
        
        column_error        : 'img/column_error.gif',
        
        check_constraint    : 'img/check_constraint.gif',
        
        check_constraint_error   : 'img/check_constraint_error.gif',
        
        query               : 'img/query.gif',
        
        query_error         : 'img/query_error.gif',

        query_warning         : 'img/query_warning.gif',

        query_info         : 'img/query_info.gif',
        
		source              : 'img/root_node.gif',
		
		root				: 'img/base.gif',

		folder			    : 'img/folder.gif',
		
		folder_warning      : 'img/folder_warning.gif',
		
		folder_error        : 'img/folder_error.gif',
		
		folder_info         : 'img/folder_info.gif',

		folderOpen	: 'img/folderopen.gif',

		database            : 'img/source_database.gif',
		
		database_warning    : 'img/source_database_warning.gif',
		
		database_error    : 'img/source_database_error.gif',
		
		database_info    : 'img/source_database_info.gif',
		
		table               : 'img/source_table.gif',
		
		table_warning       : 'img/source_table_warning.gif',
		
		table_error       : 'img/source_table_error.gif',
		
		table_info       : 'img/source_table_info.gif',
		
		index               : 'img/source_index.gif',
		
		index_warning       : 'img/source_index_warning.gif',
		
		index_error       : 'img/source_index_error.gif',
		
		index_info       : 'img/source_index_info.gif',
		
		empty				: 'img/empty.gif',

		nlPlus			: 'img/nolines_plus.gif',

		nlMinus			: 'img/nolines_minus.gif'
		
};