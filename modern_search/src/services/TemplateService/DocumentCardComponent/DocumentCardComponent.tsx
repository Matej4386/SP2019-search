import * as React from "react";
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardPreview,
    DocumentCardTitle,
    IDocumentCardPreviewProps,
    DocumentCardActions,
    DocumentCardType
  } from 'office-ui-fabric-react/lib/DocumentCard';
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import PreviewContainer from '../PreviewContainer/PreviewContainer';
import { PreviewType } from '../PreviewContainer/IPreviewContainerProps';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { IComponentFieldsConfiguration, TemplateService } from "../TemplateService";
import * as Handlebars from 'handlebars';
import * as documentCardLocationGetStyles from 'office-ui-fabric-react/lib/components/DocumentCard/DocumentCard.scss';
import { getTheme, mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import { classNamesFunction } from "office-ui-fabric-react/lib/Utilities";

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
    tags?:string;
    href?: string; 
    previewImage? :string;
    date?: string;
    profileImage?: string;
    previewUrl?: string;
    author?: string;
    iconSrc?: string;
    fileExtension?: string;

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

    private documentCardPreviewRef;

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
        
        if (this.state.showCallout && processedProps.previewUrl && this.props.enablePreview) {

            renderPreviewCallout = <PreviewContainer
                elementUrl={processedProps.previewUrl}
                previewImageUrl={processedProps.previewImage}
                previewType={processedProps.isVideo ? PreviewType.Video : PreviewType.Document}
                targetElement={this.documentCardPreviewRef}
                showPreview={this.state.showCallout}
                videoProps={{
                    fileExtension: processedProps.fileExtension
                }}
            />;
        }
        
        let previewProps: IDocumentCardPreviewProps = {
            previewImages: [
              {
                name: processedProps.title,
                previewImageSrc: processedProps.previewImage,
                imageFit: ImageFit.center,
                iconSrc: this.props.isVideo || !this.props.showFileIcon ? '' : processedProps.iconSrc,
                width: 318,
                height: 196,
              }
            ],
        };
        
        return <div style={{ marginBottom: '0.5rem'}} ref={(documentCardPreviewRefThis) => (this.documentCardPreviewRef = documentCardPreviewRefThis)}>
                    <DocumentCard 
                        onClick={(ev) => {
                            console.log (ev);
                            ev.preventDefault();
                            ev.stopPropagation();
                            this.setState({
                                showCallout: true
                            });
                        }}
                        type={ this.props.isCompact ? DocumentCardType.compact : DocumentCardType.normal }    
                    >
                        <DocumentCardPreview {...previewProps} />
                        <DocumentCardTitle
                            title={processedProps.title}
                            shouldTruncate={false}
                        />                                  
                        { processedProps.author ?
                            <DocumentCardActivity
                                activity={processedProps.date}
                                people={[{ name: processedProps.author, profileImageSrc: processedProps.profileImage}]}
                            /> : null 
                        }
                        <DocumentCardActions
                            actions={
                                [
                                    {
                                        iconProps: { iconName: 'Share' },
                                        onClick: (ev: any) => {
                                            var redirectWindow = window.open(processedProps.href, '_blank');
                                            redirectWindow.location;
                                            ev.preventDefault();
                                            ev.stopPropagation();
                                        },
                                        ariaLabel: 'share action'
                                    }
                                ]
                            }
                        />
                    </DocumentCard>
                    {renderPreviewCallout}
                </div>;
    }
}