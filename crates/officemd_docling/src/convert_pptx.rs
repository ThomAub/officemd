//! PPTX-specific conversion from `OoxmlDocument` IR to Docling structures.

use officemd_core::ir::OoxmlDocument;

use crate::builder::DoclingDocumentBuilder;
use crate::convert::{convert_blocks, convert_comments, flatten_paragraph};
use crate::model::{
    ContentLayer, DocItemLabel, GroupItem, GroupLabel, PageItem, RefItem, Size, TextItem,
};

pub fn convert_pptx(doc: &OoxmlDocument, builder: &mut DoclingDocumentBuilder) {
    for slide in &doc.slides {
        // Each slide becomes a Section group + a page entry.
        let slide_name = match &slide.title {
            Some(title) => format!("Slide {} - {}", slide.number, title),
            None => format!("Slide {}", slide.number),
        };

        let slide_ref = builder.add_group(GroupItem {
            self_ref: String::new(),
            parent: Some(RefItem::new("#/body")),
            children: vec![],
            content_layer: ContentLayer::Body,
            name: slide_name,
            label: GroupLabel::Section,
        });
        builder.add_child_to_parent("#/body", &slide_ref);

        // Page entry for the slide.
        builder.add_page(
            slide.number,
            PageItem {
                size: Size {
                    width: 0.0,
                    height: 0.0,
                },
                page_no: slide.number,
            },
        );

        // Title as a Title text item.
        if let Some(title) = &slide.title {
            let text_ref = builder.add_text(TextItem {
                self_ref: String::new(),
                parent: Some(RefItem::new(&slide_ref)),
                children: vec![],
                content_layer: ContentLayer::Body,
                label: DocItemLabel::Title,
                prov: vec![],
                orig: title.clone(),
                text: title.clone(),
                level: None,
                enumerated: None,
                marker: None,
            });
            builder.add_child_to_parent(&slide_ref, &text_ref);
        }

        // Body blocks.
        convert_blocks(&slide.blocks, &slide_ref, ContentLayer::Body, builder);

        // Speaker notes.
        if let Some(notes) = &slide.notes {
            for paragraph in notes {
                let text = flatten_paragraph(paragraph);
                let text_ref = builder.add_text(TextItem {
                    self_ref: String::new(),
                    parent: Some(RefItem::new(&slide_ref)),
                    children: vec![],
                    content_layer: ContentLayer::Body,
                    label: DocItemLabel::Text,
                    prov: vec![],
                    orig: text.clone(),
                    text,
                    level: None,
                    enumerated: None,
                    marker: None,
                });
                builder.add_child_to_parent(&slide_ref, &text_ref);
            }
        }

        // Comments.
        convert_comments(&slide.comments, &slide_ref, ContentLayer::Body, builder);
    }
}
