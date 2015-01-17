/* 
JSONScript VB6 script interpreter example
26/April/2009 by Mike Glaser
*/

// main program command array
[{
    // Shows an alert with a cat's color based on a lookup
    'alert': {
        title: 'Cat Color: ',
        message: [
        // this is the command array that makes up the message text
        {
            // set the "snowycolor" variable to a user defined value
            'set': {
                name: 'snowycolor',
                value: {
                    // get value in input box
                    'input': {
                        'title': 'Snowy the the cats\' color',
                        'prompt': 'What\'s Snowy\'s color:',
                        'default': 'White'
                    }
                }
            }
        }, {
            // check if value was "white", if so set var to "Whiter than Snow"
            'set': {
                name: 'snowycolor',
                value: {
                    // if snowycolor = 'white' then return 'Whiter than Snow'
                    'if': {
                        'value1': {
                            'get': {
                                name: 'snowycolor'
                            }
                        },
                        'value2': 'White',
                        // types of compare: eq, gt, lt, gte, lte
                        'type': 'eq',
                        'true': 'Whiter than Snow',
                        'false': {
                            'get': {
                                name: 'snowycolor'
                            }
                        }
                    }
                }
            }
        }, 
		{
            // lookup the cat's name
            'switch': {
                'case': {
                    'input': {
                        'title': 'Please Enter a cat name',
                        'prompt': 'Cat Name (Sly, Ebony, Karma, Snowy):',
                        'default': 'Snowy'
                    }
                },
                // default case
                'default': 'Unknown Color',
                // cases
                items: [{
                    'case': 'Sly',
                    'return': 'Gray'
                }, {
                    'case': 'Ebony',
                    'return': 'Black'
                }, {
                    'case': 'Karma',
                    'return': 'Ginger'
                }, {
                    'case': 'Snowy',
                    // example of string concatonation
                    'return': [' Snowy\'s Color is: ', {
                        'get': {
                            'name': 'snowycolor',
                            'default': 'None Given'
                        }
                    }]
                }]
            }
        }]
        // Program command array finished for 'message' variable, so now alert will display.
    }
}, 
// This is the string to return to the command interpreter
'The Program Ran OK']
