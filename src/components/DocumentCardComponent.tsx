import * as React from 'react';
import { IDocumentCardPreviewProps, DocumentCard, DocumentCardPreview, DocumentCardTitle, DocumentCardActivity, DocumentCardType, DocumentCardLocation } from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import PreviewContainer from '../controls/PreviewContainer/PreviewContainer';
import { PreviewType } from '../controls/PreviewContainer/IPreviewContainerProps';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TemplateService } from '../services/TemplateService/TemplateService';
import { trimStart } from '@microsoft/sp-lodash-subset';
import { getFileTypeIconProps, FileIconType } from '@uifabric/file-type-icons';
import { GlobalSettings } from 'office-ui-fabric-react/lib/Utilities'; // has to be present
import { BaseWebComponent } from './BaseWebComponent';
import * as ReactDOM from 'react-dom';
let globalSettings = (window as any).__globalSettings__;
import * as DOMPurify from 'dompurify';

/**
 * Document card props. These properties are retrieved from the web component attributes. They must be camel case.
 * (ex: a 'preview-image' HTML attribute becomes 'previewImage' prop, etc.)
 */
export interface IDocumentCardComponentProps {

    // Item context
    item?: string;

    // Fields configuration object
    fieldsConfiguration?: string;

    // Individual content properties (i.e web component attributes)
    title?: string;
    location?: string;
    tags?: string;
    href?: string;
    previewImage?: string;
    date?: string;
    profileImage?: string;
    previewUrl?: string;
    author?: string;
    iconSrc?: string;
    iconExt?: string;
    fileExtension?: string;
    UniqueId?: string;
    // Behavior properties
    enablePreview?: boolean;
    showFileIcon?: boolean;
    isVideo?: boolean;
    isCompact?: boolean;

}

export interface IDocumentCardComponentState {
    showCallout: boolean;
}

export class DocumentCardComponent extends React.Component<IDocumentCardComponentProps, IDocumentCardComponentState> {

    private documentCardPreviewRef = null; //React.createRef<HTMLDivElement>();

    public constructor(props: IDocumentCardComponentProps) {
        super(props);
        this.state = {
            showCallout: false
        };
    }

    public render() {

        let renderPreviewCallout = null;
        let processedProps: IDocumentCardComponentProps = this.props;

        if (this.props.fieldsConfiguration && this.props.item) {
            processedProps = TemplateService.processFieldsConfiguration<IDocumentCardComponentProps>(this.props.fieldsConfiguration, this.props.item);
        }
        if (this.documentCardPreviewRef !== null && this.state.showCallout && (processedProps.previewUrl || processedProps.previewImage) && this.props.enablePreview) {
            renderPreviewCallout = <PreviewContainer
                elementUrl={processedProps.previewUrl ? processedProps.previewUrl : processedProps.previewImage}
                previewImageUrl={processedProps.previewImage ? processedProps.previewImage : processedProps.previewUrl}
                previewType={processedProps.isVideo ? PreviewType.Video : PreviewType.Document}
                targetElement={this.documentCardPreviewRef}
                showPreview={this.state.showCallout}
                videoProps={{
                    fileExtension: processedProps.fileExtension
                }}
            />;
        }

        let iconSrc = processedProps.iconSrc;
        if (!iconSrc) {
            let iconProps;
            // same code as in IconComponent.tsx
            if (processedProps.iconExt) {
                if (processedProps.iconExt == 'IsListItem') {
                    iconProps = getFileTypeIconProps({ type: FileIconType.listItem, size: 32, imageFileType: 'png' });
                } else if (processedProps.iconExt == 'IsContainer') {
                    iconProps = getFileTypeIconProps({ type: FileIconType.folder, size: 32, imageFileType: 'png' });
                } else {
                    iconProps = getFileTypeIconProps({ extension: processedProps.iconExt, size: 32, imageFileType: 'png' });
                }
            } else {
                const fileExtension = processedProps.fileExtension ? trimStart(processedProps.fileExtension.trim(), '.') : null;
                iconProps = getFileTypeIconProps({ extension: fileExtension, size: 32, imageFileType: 'png' });
            }

            if (globalSettings.icons[iconProps.iconName] && this.props.showFileIcon) {
                iconSrc = globalSettings.icons[iconProps.iconName].code.props.src;
            }
        }

        let previewProps: IDocumentCardPreviewProps = {
            previewImages: [
                {
                    name: processedProps.title,
                    previewImageSrc: processedProps.previewImage,
                    imageFit: ImageFit.center,
                    iconSrc: iconSrc,
                    width: this.props.isCompact ? 144 : 318,
                    height: this.props.isCompact ? 106 : 196
                }
            ],
        };

        const playButtonStyles: React.CSSProperties = {
            color: '#fff',
            padding: '15px',
            backgroundColor: 'gray',
            left: '50%',
            top: '50%',
            transform: 'translate(-50%, -50%)',
            position: 'absolute',
            zIndex: 1,
            opacity: 0.9,
            borderRadius: '50%',
            borderColor: '#fff',
            borderWidth: 4,
            borderStyle: 'solid',
            display: 'flex',
        };
        return <div ref={(documentCardPreviewRefThis) => (this.documentCardPreviewRef = documentCardPreviewRefThis)}>
            <DocumentCard
                onClick={() => {
                    this.setState({
                        showCallout: true
                    });
                }}
                type={this.props.isCompact ? DocumentCardType.compact : DocumentCardType.normal}
            >
                <div ref={this.documentCardPreviewRef} style={{ position: 'relative', height: '100%' }}>
                    {this.props.isVideo ?
                        <div style={playButtonStyles}>
                            <i className='ms-Icon ms-Icon--Play ms-font-xl' aria-hidden='true'></i>
                        </div> : null
                    }
                    <DocumentCardPreview {...previewProps} />
                </div>
                <div>
                    {processedProps.location && !this.props.isCompact ?
                        <div style={{paddingLeft:'0.5rem'}} dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(processedProps.location) }}></div> : null
                    }
                    <Link
                        href={processedProps.href} target='_blank' data-interception='off' 
                    >
                        <DocumentCardTitle
                            title={processedProps.title}
                            shouldTruncate={false}
                        />
                    </Link>
                    {processedProps.author ?
                        <DocumentCardActivity
                            activity={processedProps.date}
                            people={[{ name: processedProps.author, profileImageSrc: processedProps.profileImage }]}
                        /> : null
                    }
                </div>
            </DocumentCard>
            {renderPreviewCallout}
        </div>;
    }
}

export class DocumentCardWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public connectedCallback() {

        let props = this.resolveAttributes();
        const documentCarditem = <DocumentCardComponent {...props} />;
        ReactDOM.render(documentCarditem, this);
    }
}

export class VideoCardWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public connectedCallback() {

        // Get all custom element attributes
        let props = this.resolveAttributes();

        // Add video props
        props.isVideo = true;

        const documentCarditem = <DocumentCardComponent {...props} />;
        ReactDOM.render(documentCarditem, this);
    }
}
