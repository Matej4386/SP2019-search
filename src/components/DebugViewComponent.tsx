import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { BaseWebComponent } from './BaseWebComponent';
import AceEditor from 'react-ace';

export interface IDebugViewProps {

    /**
     * The debug content to display
     */
    content?: string;
}

export interface IDebugViewState {
}

export default class DebugView extends React.Component<IDebugViewProps, IDebugViewState> {
    
    public render() {
        return <AceEditor
            width={ '100%' }
            mode={ 'json' }
            theme='textmate'
            enableLiveAutocompletion={ true }
            showPrintMargin={ false }
            showGutter= { true }            
            value={ this.props.content }
            highlightActiveLine={ true }
            readOnly={ true }
            editorProps={
                {
                    $blockScrolling: Infinity,
                }
            }					
            name='CodeView'
        />;
    }
}

export class DebugViewWebComponent extends BaseWebComponent {
   
    public constructor() {
        super(); 
    }
 
    public async connectedCallback() {
 
       // Reuse the 'brace' imports from the PnP control instead of reference them explicitly in the debug view
       await import(
          /* webpackChunkName: 'debug-view' */
          '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
       );
 
       let props = this.resolveAttributes();
       const debugView = <DebugView {...props}/>;
       ReactDOM.render(debugView, this);
    }    
}