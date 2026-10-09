//! Slide masters, slide layouts and themes as editable parts.
//!
//! A master or a layout owns a shape tree and a background as a slide does,
//! so [`PartRef`] addresses all three for shape and background edits. Layout
//! and master edits stay in their records and theme edits in
//! `theme_edits`, and staging the package writes them back.

use std::collections::HashSet;
use std::io::Cursor;

use oxml_drawing::color::ColorChoice;
use oxml_drawing::fill::Fill;
use oxml_drawing::text::{CT_TextListStyle, CT_TextParagraphProperties};
use oxml_drawing::theme::{CT_OfficeStyleSheet, ThemeTypeface};
use oxml_opc::relationship::{Relationship, rel_types};
use oxml_opc::{OpcPackage, Relationships, content_types};
use rpptx_oxml::namespace::R_NS;
use rpptx_oxml::relmap::{relationship_ids, rewrite_exact_rel_ids};
use rpptx_oxml::slide_parts::{
    BackgroundRendering, CT_CommonSlideData, CT_MasterTextStyles, CT_SlideLayout, CT_SlideMaster,
};

use crate::{
    Error, LayoutRecord, MediaStore, Presentation, Result, RgbColor, ShapeMut, ShapeRef, ShapesMut,
    add_image_relationship, collect_subtree_ids, detach_connectors, group_at_mut,
    invalid_presentation_mutation, invalid_shape_mutation, locate_shape_hyperlink,
    next_numbered_part_number, picture_dimensions, prune_unreachable_parts, related_internal_part,
    relationship_is_external, relative_part_target, required_part, rewrite_shape_hyperlink,
    shape_id_count, shape_kind, shape_mut, shape_ref,
};

/// A slide, slide layout or slide master, by zero-based index: the parts
/// that own a shape tree and a background.
///
/// Layout indices follow [`Presentation::layout_name`] and master indices
/// `p:sldMasterIdLst` order.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum PartRef {
    Slide(usize),
    Layout(usize),
    Master(usize),
}

/// One of a master's three text styles in `p:txStyles`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MasterTextStyle {
    /// `p:titleStyle`, the title placeholders.
    Title,
    /// `p:bodyStyle`, the body, subtitle and content placeholders.
    Body,
    /// `p:otherStyle`, text boxes and shapes that are not placeholders.
    Other,
}

/// A theme's heading or body fonts.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ThemeFontRole {
    /// The heading fonts, `+mj-lt` and its peers.
    Major,
    /// The body fonts, `+mn-lt` and its peers.
    Minor,
}

/// One script of a theme font collection.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ThemeFontScript {
    Latin,
    EastAsian,
    ComplexScript,
}

/// The theme colour slots in scheme order, as [`crate::Theme::colors`] names them.
const THEME_COLOR_SLOTS: [&str; 12] = [
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// The smallest `p:sldLayoutId` and `p:sldMasterId` value, `ST_SlideLayoutId`.
const MIN_LAYOUT_ID: u32 = 2_147_483_648;

fn malformed(part_name: &str, error: impl ToString) -> Error {
    Error::MalformedPart {
        part_name: part_name.to_owned(),
        message: error.to_string(),
    }
}

fn unknown_master(index: usize, master_count: usize) -> Error {
    Error::UnknownMasterIndex {
        index,
        master_count,
    }
}

fn unknown_layout(index: usize, layout_count: usize) -> Error {
    Error::UnknownLayoutIndex {
        index,
        layout_count,
    }
}

/// One level of a text list style, by its one-based number.
fn level_slot(
    list: &mut CT_TextListStyle,
    level: usize,
) -> &mut Option<CT_TextParagraphProperties> {
    match level {
        1 => &mut list.level1,
        2 => &mut list.level2,
        3 => &mut list.level3,
        4 => &mut list.level4,
        5 => &mut list.level5,
        6 => &mut list.level6,
        7 => &mut list.level7,
        8 => &mut list.level8,
        _ => &mut list.level9,
    }
}

fn check_text_style_level(level: usize) -> Result<()> {
    if (1..=9).contains(&level) {
        Ok(())
    } else {
        Err(invalid_presentation_mutation(
            "edit master text style",
            format!("text style level {level} is outside 1 to 9"),
        ))
    }
}

fn style_list(styles: &CT_MasterTextStyles, style: MasterTextStyle) -> &CT_TextListStyle {
    match style {
        MasterTextStyle::Title => &styles.title_style,
        MasterTextStyle::Body => &styles.body_style,
        MasterTextStyle::Other => &styles.other_style,
    }
}

/// Whether a source layout stands for a destination layout when a master
/// design is imported: the same layout type, or for custom and untyped
/// layouts the same name.
fn layouts_match(destination: &CT_SlideLayout, source: &CT_SlideLayout) -> bool {
    match (destination.layout_type(), source.layout_type()) {
        (Some(destination), Some(source)) if destination != "cust" && source != "cust" => {
            destination == source
        }
        _ => {
            destination.common_slide_data.name.is_some()
                && destination.common_slide_data.name == source.common_slide_data.name
        }
    }
}

/// The internal parts one part's relationships target, except those of
/// `kept` types, as pruning candidates once the relationships go.
fn internal_targets(
    package: &OpcPackage,
    part_name: &str,
    kept: &[&str],
    candidates: &mut HashSet<String>,
) {
    if let Some(relationships) = package.get_part_rels(part_name) {
        candidates.extend(
            relationships
                .items
                .iter()
                .filter(|relationship| {
                    !relationship_is_external(relationship)
                        && !kept.contains(&relationship.rel_type.as_str())
                })
                .map(|relationship| {
                    OpcPackage::resolve_rel_target(part_name, &relationship.target)
                }),
        );
    }
}

/// Copies the relationships of a source part that a destination part takes
/// over with its XML. Images are added to the destination's media,
/// external targets are kept, and `skipped` types, which the destination
/// supplies itself, are left out. With `fresh_ids`, each copy gets the next
/// free id of `destination` and the returned map renames the source id,
/// otherwise it keeps its id. Any other internal relationship is refused,
/// because the part it targets would be shared or lost.
#[allow(clippy::too_many_arguments)]
fn copy_relationships(
    source_package: &OpcPackage,
    source_part: &str,
    package: &mut OpcPackage,
    media_store: &mut MediaStore,
    destination_part: &str,
    skipped: &[&str],
    fresh_ids: bool,
    destination: &mut Relationships,
) -> Result<std::collections::HashMap<String, String>> {
    let mut renamed = std::collections::HashMap::new();
    let Some(source) = source_package.get_part_rels(source_part) else {
        return Ok(renamed);
    };
    for relationship in &source.items {
        if skipped.contains(&relationship.rel_type.as_str()) {
            continue;
        }
        let target = if relationship_is_external(relationship) {
            relationship.target.clone()
        } else if relationship.rel_type == rel_types::IMAGE {
            let image_part = OpcPackage::resolve_rel_target(source_part, &relationship.target);
            let bytes = required_part(source_package, &image_part)?;
            let media_part = media_store.insert(package, bytes, &image_part);
            relative_part_target(destination_part, &media_part)
        } else {
            return Err(invalid_presentation_mutation(
                "copy a master, layout or theme",
                format!(
                    "{source_part} has a {} relationship that rpptx cannot copy, only pictures and external links",
                    relationship.rel_type
                ),
            ));
        };
        let id = if fresh_ids {
            destination.add(&relationship.rel_type, &target)
        } else {
            destination.add_with_id(&relationship.id, &relationship.rel_type, &target);
            relationship.id.clone()
        };
        if let Some(added) = destination.items.iter_mut().find(|added| added.id == id) {
            added.target_mode = relationship.target_mode.clone();
        }
        renamed.insert(relationship.id.clone(), id);
    }
    Ok(renamed)
}

impl Presentation {
    /// Returns the part name of a slide, layout or master.
    pub(crate) fn part_name(&self, part: PartRef) -> Result<&str> {
        match part {
            PartRef::Slide(index) => self
                .slides
                .get(index)
                .map(|record| record.part_name.as_str())
                .ok_or(Error::UnknownSlideIndex {
                    index,
                    slide_count: self.slides.len(),
                }),
            PartRef::Layout(index) => self
                .layouts
                .get(index)
                .map(|record| record.part_name.as_str())
                .ok_or_else(|| unknown_layout(index, self.layouts.len())),
            PartRef::Master(index) => self
                .masters
                .get(index)
                .map(|record| record.part_name.as_str())
                .ok_or_else(|| unknown_master(index, self.masters.len())),
        }
    }

    /// Returns one parsed master, or the error that kept it from parsing.
    pub(crate) fn master_model(&self, master_index: usize) -> Result<&CT_SlideMaster> {
        let record = self
            .masters
            .get(master_index)
            .ok_or_else(|| unknown_master(master_index, self.masters.len()))?;
        record
            .master
            .as_ref()
            .map_err(|message| malformed(&record.part_name, message))
    }

    /// Returns one master for an edit that saving writes back.
    fn master_model_mut(&mut self, master_index: usize) -> Result<&mut CT_SlideMaster> {
        let master_count = self.masters.len();
        let record = self
            .masters
            .get_mut(master_index)
            .ok_or_else(|| unknown_master(master_index, master_count))?;
        record.dirty |= record.master.is_ok();
        match &mut record.master {
            Ok(master) => Ok(master),
            Err(message) => Err(malformed(&record.part_name, message.clone())),
        }
    }

    /// Returns the common slide data, shapes and background, of a part.
    pub(crate) fn common_data(&self, part: PartRef) -> Result<&CT_CommonSlideData> {
        match part {
            PartRef::Slide(index) => {
                self.part_name(part)?;
                Ok(&self.slides[index].slide.common_slide_data)
            }
            PartRef::Layout(index) => {
                self.part_name(part)?;
                Ok(&self.layouts[index].layout.common_slide_data)
            }
            PartRef::Master(index) => Ok(&self.master_model(index)?.common_slide_data),
        }
    }

    /// Returns the common slide data of a part for an edit that saving
    /// writes back.
    pub(crate) fn common_data_mut(&mut self, part: PartRef) -> Result<&mut CT_CommonSlideData> {
        self.part_name(part)?;
        match part {
            PartRef::Slide(index) => Ok(&mut self.slides[index].slide.common_slide_data),
            PartRef::Layout(index) => {
                let record = &mut self.layouts[index];
                record.dirty = true;
                Ok(&mut record.layout.common_slide_data)
            }
            PartRef::Master(index) => Ok(&mut self.master_model_mut(index)?.common_slide_data),
        }
    }

    /// Serialises the current model of a part.
    fn part_xml(&self, part: PartRef) -> Result<Vec<u8>> {
        let part_name = self.part_name(part)?;
        match part {
            PartRef::Slide(index) => self.slides[index].slide.to_xml(),
            PartRef::Layout(index) => self.layouts[index].layout.to_xml(),
            PartRef::Master(index) => self.master_model(index)?.to_xml(),
        }
        .map_err(|error| malformed(part_name, error))
    }

    fn part_relationship_ids(&self, part: PartRef) -> Result<HashSet<String>> {
        let part_name = self.part_name(part)?;
        relationship_ids(&self.part_xml(part)?)
            .map(|ids| ids.into_iter().collect())
            .map_err(|error| malformed(part_name, error))
    }

    /// Drops the relationships a part referenced in `before` and no longer
    /// references, and the package parts only they reached.
    fn release_relationships(&mut self, part: PartRef, before: &HashSet<String>) -> Result<()> {
        let after = self.part_relationship_ids(part)?;
        let part_name = self.part_name(part)?.to_owned();
        let mut candidates = HashSet::new();
        if let Some(relationships) = self.package.get_part_rels_mut(&part_name) {
            relationships.items.retain(|relationship| {
                if !before.contains(&relationship.id) || after.contains(&relationship.id) {
                    return true;
                }
                if !relationship_is_external(relationship) {
                    candidates.insert(OpcPackage::resolve_rel_target(
                        &part_name,
                        &relationship.target,
                    ));
                }
                false
            });
        }
        if !candidates.is_empty() {
            prune_unreachable_parts(&mut self.package, &candidates);
            self.media_store = MediaStore::scan(&self.package);
        }
        Ok(())
    }

    /// Returns the index of the master a layout belongs to.
    pub(crate) fn layout_master_index(&self, layout_index: usize) -> Result<usize> {
        let record = self
            .layouts
            .get(layout_index)
            .ok_or_else(|| unknown_layout(layout_index, self.layouts.len()))?;
        let master_part =
            related_internal_part(&self.package, &record.part_name, rel_types::SLIDE_MASTER)?
                .ok_or_else(|| {
                    malformed(&record.part_name, "layout has no slide master relationship")
                })?;
        self.masters
            .iter()
            .position(|master| master.part_name.eq_ignore_ascii_case(&master_part))
            .ok_or_else(|| {
                malformed(
                    &record.part_name,
                    "the layout's slide master is not in the presentation's master list",
                )
            })
    }

    /// Writes edited layouts, masters and themes into a staged package.
    pub(crate) fn write_master_parts(&self, package: &mut OpcPackage) -> Result<()> {
        for record in self.layouts.iter().filter(|record| record.dirty) {
            let xml = record
                .layout
                .to_xml()
                .map_err(|error| malformed(&record.part_name, error))?;
            package.set_part(&record.part_name, xml);
        }
        for record in self.masters.iter().filter(|record| record.dirty) {
            if let Ok(master) = &record.master {
                let xml = master
                    .to_xml()
                    .map_err(|error| malformed(&record.part_name, error))?;
                package.set_part(&record.part_name, xml);
            }
        }
        for (part_name, theme) in &self.theme_edits {
            let xml = theme
                .to_xml()
                .map_err(|error| malformed(part_name, error))?;
            package.set_part(part_name, xml);
        }
        Ok(())
    }

    /// Returns the index of the master a layout belongs to, `None` for an
    /// unknown layout.
    pub fn layout_master(&self, layout_index: usize) -> Option<usize> {
        self.layout_master_index(layout_index).ok()
    }

    /// Returns the zero-based indices of the slides that use one layout.
    pub fn layout_slides(&self, layout_index: usize) -> Vec<usize> {
        (0..self.slides.len())
            .filter(|&slide_index| self.slide_layout_index(slide_index) == Some(layout_index))
            .collect()
    }

    /// Iterates the immediate shapes of a slide, layout or master in
    /// z-order.
    pub fn part_shapes(
        &self,
        part: PartRef,
    ) -> Result<impl ExactSizeIterator<Item = ShapeRef<'_>>> {
        Ok(self
            .common_data(part)?
            .shape_tree
            .children
            .iter()
            .map(shape_ref))
    }

    /// Returns one immediate shape of a slide, layout or master.
    pub fn part_shape(&self, part: PartRef, index: usize) -> Option<ShapeRef<'_>> {
        self.common_data(part)
            .ok()?
            .shape_tree
            .children
            .get(index)
            .map(shape_ref)
    }

    /// Returns one immediate shape of a slide, layout or master for in-place
    /// mutation, which saving writes back.
    pub fn part_shape_mut(&mut self, part: PartRef, index: usize) -> Option<ShapeMut<'_>> {
        self.common_data_mut(part)
            .ok()?
            .shape_tree
            .children
            .get_mut(index)
            .map(shape_mut)
    }

    /// Returns the shapes of a slide, layout or master, or of one group in
    /// it, for adding shapes, as [`Self::shapes_mut`] does for slides.
    ///
    /// A shape added to a master shows on every slide of its layouts that
    /// show master shapes, and one added to a layout on every slide that
    /// uses it. [`crate::ShapesMut::add_picture`] adds the image
    /// relationship to that part.
    pub fn part_shapes_mut(&mut self, part: PartRef, group: &[usize]) -> Option<ShapesMut<'_>> {
        let tree = &mut self.common_data_mut(part).ok()?.shape_tree;
        if !group.is_empty() {
            group_at_mut(&mut tree.children, group)?;
        }
        Some(ShapesMut {
            presentation: self,
            part,
            group: group.to_vec(),
        })
    }

    /// Removes one immediate shape of a slide, layout or master, and the
    /// relationships and package parts only it used.
    ///
    /// A slide shape is removed as [`Self::remove_shape`] removes it.
    pub fn remove_part_shape(&mut self, part: PartRef, shape_index: usize) -> Result<()> {
        const OPERATION: &str = "remove shape";
        if let PartRef::Slide(slide_index) = part {
            return self.remove_shape(slide_index, shape_index);
        }
        let mut staged = self.clone();
        let before = staged.part_relationship_ids(part)?;
        let tree = &mut staged.common_data_mut(part)?.shape_tree;
        let child = tree.children.get(shape_index).ok_or_else(|| {
            invalid_shape_mutation(
                OPERATION,
                format!(
                    "shape index {shape_index} is out of range for {} shapes",
                    tree.children.len()
                ),
            )
        })?;
        child
            .non_visual_id()
            .ok_or_else(|| Error::UnsupportedShapeMutation {
                operation: OPERATION,
                shape_kind: shape_kind(child),
            })?;
        let mut removed_ids = HashSet::new();
        let mut media_ids = Vec::new();
        collect_subtree_ids(
            std::slice::from_ref(child),
            &mut removed_ids,
            &mut media_ids,
        );
        tree.children.remove(shape_index);
        detach_connectors(&mut tree.children, &removed_ids);
        staged.release_relationships(part, &before)?;
        self.commit_candidate(staged)
    }

    /// Moves one immediate shape of a slide, layout or master to z-order
    /// index `to_index`, where later shapes draw on top.
    pub fn move_part_shape(
        &mut self,
        part: PartRef,
        from_index: usize,
        to_index: usize,
    ) -> Result<()> {
        const OPERATION: &str = "move shape";
        if let PartRef::Slide(slide_index) = part {
            return self.move_shape(slide_index, from_index, to_index);
        }
        let tree = &mut self.common_data_mut(part)?.shape_tree;
        let count = tree.children.len();
        if let Some(index) = [from_index, to_index]
            .into_iter()
            .find(|index| *index >= count)
        {
            return Err(invalid_shape_mutation(
                OPERATION,
                format!("shape index {index} is out of range for {count} shapes"),
            ));
        }
        tree.move_child(from_index, to_index)
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))
    }

    /// Returns the external click hyperlink of one shape of a slide, layout
    /// or master, found by its non-visual id, as
    /// [`Self::shape_hyperlink_address`] does for slides.
    pub fn part_shape_hyperlink_address(
        &self,
        part: PartRef,
        shape_id: u32,
    ) -> Result<Option<&str>> {
        if let PartRef::Slide(slide_index) = part {
            return self.shape_hyperlink_address(slide_index, shape_id);
        }
        let (_, location) = self.locate_part_click(part, shape_id)?;
        let part_name = self.part_name(part)?;
        Ok(location.relationship_id.as_deref().and_then(|id| {
            self.package
                .get_part_rels(part_name)?
                .get_by_id(id)
                .map(|relationship| relationship.target.as_str())
        }))
    }

    /// Points the click action of one shape of a slide, layout or master at
    /// an external `address`, or removes it with `None`, as
    /// [`Self::set_shape_hyperlink`] does for slides. A clickable logo on a
    /// master links from every slide that shows it.
    pub fn set_part_shape_hyperlink(
        &mut self,
        part: PartRef,
        shape_id: u32,
        address: Option<&str>,
    ) -> Result<()> {
        const OPERATION: &str = "set shape hyperlink";
        if let PartRef::Slide(slide_index) = part {
            return self.set_shape_hyperlink(slide_index, shape_id, address);
        }
        if address.is_some_and(|address| {
            address.is_empty()
                || address.chars().any(|character| {
                    character.is_control() || matches!(character, '\u{FFFE}' | '\u{FFFF}')
                })
        }) {
            return Err(invalid_shape_mutation(
                OPERATION,
                "a hyperlink address must be non-empty text without control characters",
            ));
        }
        let (xml, location) = self.locate_part_click(part, shape_id)?;
        let part_name = self.part_name(part)?.to_owned();
        let mut relationships = self
            .package
            .get_part_rels(&part_name)
            .cloned()
            .unwrap_or_default();
        let link = match address {
            None if location.click.is_none() => return Ok(()),
            None => None,
            Some(address) => Some(
                relationships
                    .items
                    .iter()
                    .find(|relationship| {
                        relationship.rel_type == rel_types::HYPERLINK
                            && relationship_is_external(relationship)
                            && relationship.target == address
                    })
                    .map(|relationship| relationship.id.clone())
                    .unwrap_or_else(|| relationships.add_external(rel_types::HYPERLINK, address)),
            ),
        };
        let before = self.part_relationship_ids(part)?;
        let rewritten =
            rewrite_shape_hyperlink(&xml, &location, link.as_deref().map(|id| (id, None)))?;
        match part {
            PartRef::Layout(index) => {
                let record = &mut self.layouts[index];
                record.layout = CT_SlideLayout::from_xml(&rewritten)
                    .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
                record.dirty = true;
            }
            PartRef::Master(index) => {
                let master = CT_SlideMaster::from_xml(&rewritten)
                    .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?;
                let record = &mut self.masters[index];
                record.master = Ok(master);
                record.dirty = true;
            }
            PartRef::Slide(_) => unreachable!("slides are handled above"),
        }
        let after = self.part_relationship_ids(part)?;
        relationships.items.retain(|relationship| {
            !before.contains(&relationship.id) || after.contains(&relationship.id)
        });
        self.package.set_part_rels(&part_name, relationships);
        Ok(())
    }

    /// Serialises a layout or master and finds one shape's click action.
    fn locate_part_click(
        &self,
        part: PartRef,
        shape_id: u32,
    ) -> Result<(Vec<u8>, crate::ShapeHyperlinkLocation)> {
        const OPERATION: &str = "shape hyperlink";
        if shape_id_count(&self.common_data(part)?.shape_tree.children, shape_id) != 1 {
            return Err(invalid_shape_mutation(
                OPERATION,
                "shape id is missing or ambiguous",
            ));
        }
        let xml = self.part_xml(part)?;
        let location = locate_shape_hyperlink(&xml, shape_id)?
            .ok_or_else(|| invalid_shape_mutation(OPERATION, "shape has no p:cNvPr"))?;
        Ok((xml, location))
    }

    /// Returns the direct `p:bgPr` fill of a slide, layout or master
    /// background, `None` when it has none or follows a theme style.
    pub fn part_background_fill(&self, part: PartRef) -> Result<Option<&Fill>> {
        let Some(background) = self.common_data(part)?.background.as_ref() else {
            return Ok(None);
        };
        Ok(match background.rendering() {
            BackgroundRendering::Properties(fill) => fill.as_deref(),
            BackgroundRendering::Reference { .. } | BackgroundRendering::Unsupported(_) => None,
        })
    }

    /// Returns whether a slide, layout or master carries its own `p:bg`.
    /// Without one, a slide follows its layout and a layout its master.
    pub fn part_has_background(&self, part: PartRef) -> Result<bool> {
        Ok(self.common_data(part)?.background.is_some())
    }

    /// Sets a solid, gradient, pattern or no-fill background on a slide,
    /// layout or master. A direct-fill background keeps its other children,
    /// and a theme style reference is replaced. A picture fill needs its
    /// image relationship, so it goes through
    /// [`Self::set_part_picture_background`].
    pub fn set_part_background(&mut self, part: PartRef, fill: Fill) -> Result<()> {
        const OPERATION: &str = "set background";
        if matches!(fill, Fill::Blip(_)) {
            return Err(invalid_presentation_mutation(
                OPERATION,
                "a picture background needs its image relationship, use set_part_picture_background"
                    .to_owned(),
            ));
        }
        let before = self.part_relationship_ids(part)?;
        self.common_data_mut(part)?
            .set_background_fill(fill)
            .map_err(|error| invalid_presentation_mutation(OPERATION, error.to_string()))?;
        self.release_relationships(part, &before)
    }

    /// Sets a picture background, stretched over the whole slide, on a
    /// slide, layout or master, adding the image part and its relationship
    /// to that part.
    pub fn set_part_picture_background(
        &mut self,
        part: PartRef,
        image_data: &[u8],
        image_filename: &str,
    ) -> Result<()> {
        const OPERATION: &str = "set picture background";
        picture_dimensions(image_data, image_filename, None, None)?;
        let part_name = self.part_name(part)?.to_owned();
        let before = self.part_relationship_ids(part)?;
        let mut package = self.package.clone();
        let mut media_store = self.media_store.clone();
        let relationship_id = add_image_relationship(
            &mut package,
            &mut media_store,
            &part_name,
            image_data,
            image_filename,
        );
        let fill = Fill::from_xml(
            format!(
                r#"<a:blipFill xmlns:a="{}" xmlns:r="{R_NS}" dpi="0" rotWithShape="1"><a:blip r:embed="{relationship_id}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"#,
                oxml_drawing::namespace::A_NS
            )
            .as_bytes(),
        )
        .map_err(|error| invalid_presentation_mutation(OPERATION, error.to_string()))?;
        let mut staged_data = self.common_data(part)?.clone();
        staged_data
            .set_background_fill(fill)
            .map_err(|error| invalid_presentation_mutation(OPERATION, error.to_string()))?;
        self.package = package;
        self.media_store = media_store;
        *self.common_data_mut(part)? = staged_data;
        self.release_relationships(part, &before)
    }

    /// Removes the background of a slide, layout or master, so a slide
    /// follows its layout and a layout its master. A master without a
    /// background shows white.
    pub fn remove_part_background(&mut self, part: PartRef) -> Result<()> {
        let before = self.part_relationship_ids(part)?;
        self.common_data_mut(part)?.background = None;
        self.release_relationships(part, &before)
    }

    /// Returns whether a slide or layout shows the shapes of the parts
    /// above it, `showMasterSp`, which hides a master logo when false.
    pub fn show_master_shapes(&self, part: PartRef) -> Result<bool> {
        self.part_name(part)?;
        match part {
            PartRef::Slide(index) => {
                Ok(self.slides[index].slide.show_master_shapes.unwrap_or(true))
            }
            PartRef::Layout(index) => Ok(self.layouts[index]
                .layout
                .show_master_shapes
                .unwrap_or(true)),
            PartRef::Master(_) => Err(invalid_presentation_mutation(
                "read show master shapes",
                "a slide master has no parent shapes to show or hide, use it on a layout or slide"
                    .to_owned(),
            )),
        }
    }

    /// Sets whether a slide or layout shows the shapes of the parts above
    /// it. Hiding them on a layout hides a master logo on every slide that
    /// uses the layout. Showing them removes the attribute.
    pub fn set_show_master_shapes(&mut self, part: PartRef, show: bool) -> Result<()> {
        self.part_name(part)?;
        let value = (!show).then_some(false);
        match part {
            PartRef::Slide(index) => self.slides[index].slide.show_master_shapes = value,
            PartRef::Layout(index) => {
                let record = &mut self.layouts[index];
                record.layout.show_master_shapes = value;
                record.dirty = true;
            }
            PartRef::Master(_) => {
                return Err(invalid_presentation_mutation(
                    "set show master shapes",
                    "a slide master has no parent shapes to show or hide, use it on a layout or slide"
                        .to_owned(),
                ));
            }
        }
        Ok(())
    }

    /// Renames one layout, its `p:cSld/@name`.
    pub fn set_layout_name(&mut self, layout_index: usize, name: &str) -> Result<()> {
        let layout_count = self.layouts.len();
        let record = self
            .layouts
            .get_mut(layout_index)
            .ok_or_else(|| unknown_layout(layout_index, layout_count))?;
        record.layout.common_slide_data.name = Some(name.to_owned());
        record.dirty = true;
        Ok(())
    }

    /// Removes one layout that no slide uses, with its relationship and
    /// list entry on its master and the parts only it used.
    ///
    /// A layout in use is refused, as python-pptx `SlideLayouts.remove`
    /// refuses it, and so is the last layout of a master. Layout indices
    /// after it move down by one.
    pub fn remove_layout(&mut self, layout_index: usize) -> Result<()> {
        const OPERATION: &str = "remove slide layout";
        let master_index = self.layout_master_index(layout_index)?;
        let name = self.layout_name(layout_index).unwrap_or("").to_owned();
        let used = self.layout_slides(layout_index);
        if !used.is_empty() {
            return Err(invalid_presentation_mutation(
                OPERATION,
                format!(
                    "layout {name:?} is used by slides {used:?}, give them another layout with set_slide_layout first"
                ),
            ));
        }
        if self
            .master_layouts(master_index)
            .is_some_and(|layouts| layouts.len() <= 1)
        {
            return Err(invalid_presentation_mutation(
                OPERATION,
                format!("layout {name:?} is the only layout of its slide master"),
            ));
        }
        let master_part = self.masters[master_index].part_name.clone();
        let layout_part = self.layouts[layout_index].part_name.clone();
        let mut staged = self.clone();
        let mut removed = Vec::new();
        if let Some(relationships) = staged.package.get_part_rels_mut(&master_part) {
            relationships.items.retain(|relationship| {
                let targets_layout = relationship.rel_type == rel_types::SLIDE_LAYOUT
                    && !relationship_is_external(relationship)
                    && OpcPackage::resolve_rel_target(&master_part, &relationship.target)
                        .eq_ignore_ascii_case(&layout_part);
                if targets_layout {
                    removed.push(relationship.id.clone());
                }
                !targets_layout
            });
        }
        let master = staged.master_model_mut(master_index)?;
        let entries = master
            .slide_layout_ids()
            .map_err(|error| malformed(&master_part, error))?
            .into_iter()
            .filter(|(_, relationship_id)| !removed.contains(relationship_id))
            .collect::<Vec<_>>();
        master
            .set_slide_layout_ids(&entries)
            .map_err(|error| malformed(&master_part, error))?;
        let mut candidates = HashSet::new();
        internal_targets(
            &staged.package,
            &layout_part,
            &[rel_types::SLIDE_MASTER],
            &mut candidates,
        );
        staged.package.remove_part(&layout_part);
        staged.package.remove_part_rels(&layout_part);
        staged.package.content_types.remove_override(&layout_part);
        staged.layouts.remove(layout_index);
        prune_unreachable_parts(&mut staged.package, &candidates);
        staged.media_store = MediaStore::scan(&staged.package);
        self.commit_candidate(staged)
    }

    /// Copies one layout into a new layout of the same master, placed right
    /// after it, and returns the new layout's index.
    ///
    /// The copy is named as PowerPoint names one, `1_Title Slide` for
    /// `Title Slide`, and shares the source's pictures. A layout that holds
    /// a chart, diagram or embedded object is refused, because the copy
    /// would share that part.
    pub fn duplicate_layout(&mut self, layout_index: usize) -> Result<usize> {
        const OPERATION: &str = "duplicate slide layout";
        let master_index = self.layout_master_index(layout_index)?;
        let master_part = self.masters[master_index].part_name.clone();
        let source_part = self.layouts[layout_index].part_name.clone();
        let relationships = self
            .package
            .get_part_rels(&source_part)
            .cloned()
            .unwrap_or_default();
        if let Some(relationship) = relationships.items.iter().find(|relationship| {
            !relationship_is_external(relationship)
                && ![
                    rel_types::IMAGE,
                    rel_types::SLIDE_MASTER,
                    rel_types::HYPERLINK,
                ]
                .contains(&relationship.rel_type.as_str())
        }) {
            return Err(invalid_presentation_mutation(
                OPERATION,
                format!(
                    "{source_part} has a {} relationship whose part a copy would share",
                    relationship.rel_type
                ),
            ));
        }
        let mut staged = self.clone();
        let number =
            next_numbered_part_number(&staged.package, "/ppt/slideLayouts/slideLayout", ".xml");
        let new_part = format!("/ppt/slideLayouts/slideLayout{number}.xml");
        let mut layout = staged.layouts[layout_index].layout.clone();
        let name = layout.common_slide_data.name.clone().unwrap_or_default();
        let mut prefix = 1;
        let new_name =
            loop {
                let candidate = format!("{prefix}_{name}");
                if !staged.layouts.iter().any(|record| {
                    record.layout.common_slide_data.name.as_deref() == Some(&candidate)
                }) {
                    break candidate;
                }
                prefix += 1;
            };
        layout.common_slide_data.name = Some(new_name);
        let xml = layout
            .to_xml()
            .map_err(|error| malformed(&new_part, error))?;
        staged.package.set_part(&new_part, xml);
        staged.package.set_part_rels(&new_part, relationships);
        staged
            .package
            .content_types
            .add_override(&new_part, content_types::SLIDE_LAYOUT);
        let next_id = staged.next_layout_id()?;
        let master_relationships = staged.package.get_or_create_part_rels(&master_part);
        let source_relationship = master_relationships
            .items
            .iter()
            .position(|relationship| {
                relationship.rel_type == rel_types::SLIDE_LAYOUT
                    && OpcPackage::resolve_rel_target(&master_part, &relationship.target)
                        .eq_ignore_ascii_case(&source_part)
            })
            .ok_or_else(|| {
                malformed(&master_part, "the master has no relationship to the layout")
            })?;
        let source_relationship_id = master_relationships.items[source_relationship].id.clone();
        let new_relationship_id = master_relationships.add(
            rel_types::SLIDE_LAYOUT,
            &relative_part_target(&master_part, &new_part),
        );
        let added: Relationship = master_relationships
            .items
            .pop()
            .expect("the new relationship was just added");
        master_relationships
            .items
            .insert(source_relationship + 1, added);
        let master = staged.master_model_mut(master_index)?;
        let mut entries = master
            .slide_layout_ids()
            .map_err(|error| malformed(&master_part, error))?;
        let at = entries
            .iter()
            .position(|(_, relationship_id)| *relationship_id == source_relationship_id)
            .map_or(entries.len(), |position| position + 1);
        entries.insert(at, (next_id, new_relationship_id));
        master
            .set_slide_layout_ids(&entries)
            .map_err(|error| malformed(&master_part, error))?;
        self.commit_candidate(staged)?;
        self.layouts
            .iter()
            .position(|record| record.part_name.eq_ignore_ascii_case(&new_part))
            .ok_or_else(|| malformed(&new_part, "the duplicated layout was not retained"))
    }

    /// The next unused `p:sldLayoutId` value, above every layout and master
    /// id of the presentation, as the two lists share one id space.
    fn next_layout_id(&self) -> Result<u32> {
        let mut highest = MIN_LAYOUT_ID - 1;
        for master_id in &self.presentation.slide_master_ids {
            highest = highest.max(master_id.id.unwrap_or(0));
        }
        for index in 0..self.masters.len() {
            let master = self.master_model(index)?;
            for (id, _) in master
                .slide_layout_ids()
                .map_err(|error| malformed(&self.masters[index].part_name, error))?
            {
                highest = highest.max(id);
            }
        }
        highest.checked_add(1).ok_or_else(|| {
            invalid_presentation_mutation(
                "add slide layout",
                "slide layout ids are exhausted".to_owned(),
            )
        })
    }

    /// Returns one level of a master text style, `None` when the master
    /// sets nothing for it. `level` is one-based, 1 to 9.
    pub fn master_text_style(
        &self,
        master_index: usize,
        style: MasterTextStyle,
        level: usize,
    ) -> Result<Option<&CT_TextParagraphProperties>> {
        check_text_style_level(level)?;
        Ok(self
            .master_model(master_index)?
            .text_styles
            .as_ref()
            .and_then(|styles| style_list(styles, style).level(level)))
    }

    /// Returns one level of a master text style for editing, creating it
    /// when absent. Its default run properties set the font, size and
    /// colour, and its bullet and indents the list look, of every
    /// placeholder that inherits the level. `level` is one-based, 1 to 9.
    pub fn master_text_style_mut(
        &mut self,
        master_index: usize,
        style: MasterTextStyle,
        level: usize,
    ) -> Result<&mut CT_TextParagraphProperties> {
        check_text_style_level(level)?;
        let master = self.master_model_mut(master_index)?;
        let styles = master
            .text_styles
            .get_or_insert_with(CT_MasterTextStyles::default);
        let list = match style {
            MasterTextStyle::Title => &mut styles.title_style,
            MasterTextStyle::Body => &mut styles.body_style,
            MasterTextStyle::Other => &mut styles.other_style,
        };
        Ok(level_slot(list, level).get_or_insert_with(CT_TextParagraphProperties::default))
    }

    /// Returns the theme part of one master.
    fn master_theme_part(&self, master_index: usize) -> Result<String> {
        let master_part = self.part_name(PartRef::Master(master_index))?;
        related_internal_part(&self.package, master_part, rel_types::THEME)?
            .ok_or_else(|| malformed(master_part, "slide master has no theme relationship"))
    }

    /// Returns one master's theme for an edit that saving writes back.
    fn theme_mut(&mut self, master_index: usize) -> Result<&mut CT_OfficeStyleSheet> {
        let theme_part = self.master_theme_part(master_index)?;
        if !self.theme_edits.contains_key(&theme_part) {
            let theme = CT_OfficeStyleSheet::from_xml(required_part(&self.package, &theme_part)?)
                .map_err(|error| malformed(&theme_part, error))?;
            self.theme_edits.insert(theme_part.clone(), theme);
        }
        Ok(self
            .theme_edits
            .get_mut(&theme_part)
            .expect("the theme edit was just inserted"))
    }

    /// Sets one colour of a master's theme, by its scheme slot name: `dk1`,
    /// `lt1`, `dk2`, `lt2`, `accent1` to `accent6`, `hlink` or `folHlink`.
    /// Every shape and text that uses the scheme colour follows. Masters
    /// that share the theme part share the change.
    pub fn set_theme_color(
        &mut self,
        master_index: usize,
        slot: &str,
        color: RgbColor,
    ) -> Result<()> {
        if !THEME_COLOR_SLOTS.contains(&slot) {
            return Err(invalid_presentation_mutation(
                "set theme color",
                format!(
                    "unknown theme colour {slot:?}, use one of {}",
                    THEME_COLOR_SLOTS.join(", ")
                ),
            ));
        }
        let scheme = &mut self.theme_mut(master_index)?.theme_elements.color_scheme;
        let target = match slot {
            "dk1" => &mut scheme.dark1,
            "lt1" => &mut scheme.light1,
            "dk2" => &mut scheme.dark2,
            "lt2" => &mut scheme.light2,
            "accent1" => &mut scheme.accent1,
            "accent2" => &mut scheme.accent2,
            "accent3" => &mut scheme.accent3,
            "accent4" => &mut scheme.accent4,
            "accent5" => &mut scheme.accent5,
            "accent6" => &mut scheme.accent6,
            "hlink" => &mut scheme.hyperlink,
            _ => &mut scheme.followed_hyperlink,
        };
        *target = ColorChoice::srgb(color);
        Ok(())
    }

    /// Sets one typeface of a master's theme fonts. The Latin heading font
    /// is what `+mj-lt` titles use, the Latin body font what `+mn-lt` body
    /// text uses. An empty East Asian or complex-script typeface means
    /// none, and an empty Latin one is refused.
    pub fn set_theme_font(
        &mut self,
        master_index: usize,
        role: ThemeFontRole,
        script: ThemeFontScript,
        typeface: &str,
    ) -> Result<()> {
        if script == ThemeFontScript::Latin && typeface.trim().is_empty() {
            return Err(invalid_presentation_mutation(
                "set theme font",
                "a theme's Latin font must name a typeface".to_owned(),
            ));
        }
        let scheme = &mut self.theme_mut(master_index)?.theme_elements.font_scheme;
        let collection = match role {
            ThemeFontRole::Major => &mut scheme.major_font,
            ThemeFontRole::Minor => &mut scheme.minor_font,
        };
        let face: &mut ThemeTypeface = match script {
            ThemeFontScript::Latin => &mut collection.latin,
            ThemeFontScript::EastAsian => &mut collection.east_asian,
            ThemeFontScript::ComplexScript => &mut collection.complex_script,
        };
        face.typeface = typeface.to_owned();
        Ok(())
    }

    /// Applies the theme of another presentation's first slide master to
    /// every master of this one: colours, fonts and format scheme.
    ///
    /// With `import_master`, the source master's look comes too: its
    /// shapes, such as a logo, background, text styles and header and
    /// footer settings replace each master's, and each source layout
    /// replaces the layout of the same type, or for custom layouts the same
    /// name, adding the layouts that have no counterpart. Slides keep their
    /// own content and layouts, so they take the new design, and layouts
    /// without a counterpart stay as they were. The slide size is kept.
    pub fn apply_theme(&mut self, source: &Presentation, import_master: bool) -> Result<()> {
        const OPERATION: &str = "apply theme";
        if source.masters.is_empty() {
            return Err(invalid_presentation_mutation(
                OPERATION,
                "the source presentation has no slide master".to_owned(),
            ));
        }
        let source_package = source.staged_package(false)?;
        let source_master_part = source.masters[0].part_name.clone();
        let source_theme_part =
            related_internal_part(&source_package, &source_master_part, rel_types::THEME)?
                .ok_or_else(|| {
                    malformed(
                        &source_master_part,
                        "slide master has no theme relationship",
                    )
                })?;
        let mut staged = self.clone();
        staged.replace_themes(&source_package, &source_theme_part)?;
        if import_master {
            staged.import_master_design(&source_package, &source_master_part)?;
        }
        self.commit_candidate(staged)
    }

    /// Applies a theme from the bytes of a `.pptx` or `.potx`, as
    /// [`Self::apply_theme`] does, or of a `.thmx` theme file, which
    /// applies its theme only.
    pub fn apply_theme_bytes(&mut self, bytes: &[u8], import_master: bool) -> Result<()> {
        let package = OpcPackage::from_reader(Cursor::new(bytes))?;
        if let Some(main_part) = package.main_document_part()
            && package.content_types.content_type_for(&main_part) == Some(content_types::THEME)
        {
            if import_master {
                return Err(invalid_presentation_mutation(
                    "apply theme",
                    "a .thmx theme file applies its theme only, import a master from a .pptx or .potx"
                        .to_owned(),
                ));
            }
            let mut staged = self.clone();
            staged.replace_themes(&package, &main_part)?;
            return self.commit_candidate(staged);
        }
        let source = Presentation::from_package(package)?;
        self.apply_theme(&source, import_master)
    }

    /// Replaces the theme part of every master with a source theme part
    /// and its pictures.
    fn replace_themes(&mut self, source_package: &OpcPackage, source_theme: &str) -> Result<()> {
        let xml = required_part(source_package, source_theme)?.to_vec();
        CT_OfficeStyleSheet::from_xml(&xml).map_err(|error| malformed(source_theme, error))?;
        let mut theme_parts = Vec::new();
        for master_index in 0..self.masters.len() {
            let part = self.master_theme_part(master_index)?;
            if !theme_parts.contains(&part) {
                theme_parts.push(part);
            }
        }
        let mut candidates = HashSet::new();
        for theme_part in theme_parts {
            self.theme_edits.remove(&theme_part);
            internal_targets(&self.package, &theme_part, &[], &mut candidates);
            let mut relationships = Relationships::new();
            copy_relationships(
                source_package,
                source_theme,
                &mut self.package,
                &mut self.media_store,
                &theme_part,
                &[],
                false,
                &mut relationships,
            )?;
            self.package.set_part(&theme_part, xml.clone());
            if relationships.items.is_empty() {
                self.package.remove_part_rels(&theme_part);
            } else {
                self.package.set_part_rels(&theme_part, relationships);
            }
        }
        prune_unreachable_parts(&mut self.package, &candidates);
        self.media_store = MediaStore::scan(&self.package);
        Ok(())
    }

    /// Replaces every master's design and matching layouts with a source
    /// master and its layouts, as [`Self::apply_theme`] documents.
    fn import_master_design(
        &mut self,
        source_package: &OpcPackage,
        source_master_part: &str,
    ) -> Result<()> {
        let source_master =
            CT_SlideMaster::from_xml(required_part(source_package, source_master_part)?)
                .map_err(|error| malformed(source_master_part, error))?;
        let source_master_relationships = source_package
            .get_part_rels(source_master_part)
            .cloned()
            .unwrap_or_default();
        let mut source_layouts = Vec::new();
        for (_, relationship_id) in source_master
            .slide_layout_ids()
            .map_err(|error| malformed(source_master_part, error))?
        {
            let relationship = source_master_relationships
                .get_by_id(&relationship_id)
                .ok_or_else(|| Error::MissingRelationship {
                    source_part: source_master_part.to_owned(),
                    relationship_id: relationship_id.clone(),
                })?;
            let part = OpcPackage::resolve_rel_target(source_master_part, &relationship.target);
            let layout = CT_SlideLayout::from_xml(required_part(source_package, &part)?)
                .map_err(|error| malformed(&part, error))?;
            source_layouts.push((part, layout));
        }
        let source_master_xml = source_master
            .to_xml()
            .map_err(|error| malformed(source_master_part, error))?;
        let mut next_id = self.next_layout_id()?;
        let mut candidates = HashSet::new();
        for master_index in 0..self.masters.len() {
            let master_part = self.masters[master_index].part_name.clone();
            let mut entries = self
                .master_model(master_index)?
                .slide_layout_ids()
                .map_err(|error| malformed(&master_part, error))?;
            let destination_layouts = self.master_layouts(master_index).unwrap_or_default();
            internal_targets(
                &self.package,
                &master_part,
                &[rel_types::SLIDE_LAYOUT, rel_types::THEME],
                &mut candidates,
            );
            let mut relationships = Relationships::new();
            if let Some(existing) = self.package.get_part_rels(&master_part) {
                relationships.items.extend(
                    existing
                        .items
                        .iter()
                        .filter(|relationship| {
                            [rel_types::SLIDE_LAYOUT, rel_types::THEME]
                                .contains(&relationship.rel_type.as_str())
                        })
                        .cloned(),
                );
            }
            let renamed = copy_relationships(
                source_package,
                source_master_part,
                &mut self.package,
                &mut self.media_store,
                &master_part,
                &[rel_types::SLIDE_LAYOUT, rel_types::THEME],
                true,
                &mut relationships,
            )?;
            let xml = rewrite_exact_rel_ids(&source_master_xml, &renamed)
                .map_err(|error| malformed(&master_part, error))?;
            let mut master =
                CT_SlideMaster::from_xml(&xml).map_err(|error| malformed(&master_part, error))?;

            let mut matched = HashSet::new();
            for (source_part, source_layout) in &source_layouts {
                let source_xml = source_layout
                    .to_xml()
                    .map_err(|error| malformed(source_part, error))?;
                let target = destination_layouts.iter().copied().find(|index| {
                    !matched.contains(index)
                        && layouts_match(&self.layouts[*index].layout, source_layout)
                });
                let (layout_part, mut layout_relationships) = match target {
                    Some(index) => {
                        matched.insert(index);
                        let part = self.layouts[index].part_name.clone();
                        internal_targets(
                            &self.package,
                            &part,
                            &[rel_types::SLIDE_MASTER],
                            &mut candidates,
                        );
                        let mut kept = Relationships::new();
                        if let Some(existing) = self.package.get_part_rels(&part) {
                            kept.items.extend(
                                existing
                                    .items
                                    .iter()
                                    .filter(|relationship| {
                                        relationship.rel_type == rel_types::SLIDE_MASTER
                                    })
                                    .cloned(),
                            );
                        }
                        (part, kept)
                    }
                    None => {
                        let number = next_numbered_part_number(
                            &self.package,
                            "/ppt/slideLayouts/slideLayout",
                            ".xml",
                        );
                        let part = format!("/ppt/slideLayouts/slideLayout{number}.xml");
                        let mut kept = Relationships::new();
                        kept.add(
                            rel_types::SLIDE_MASTER,
                            &relative_part_target(&part, &master_part),
                        );
                        let relationship_id = relationships.add(
                            rel_types::SLIDE_LAYOUT,
                            &relative_part_target(&master_part, &part),
                        );
                        entries.push((next_id, relationship_id));
                        next_id = next_id.checked_add(1).ok_or_else(|| {
                            invalid_presentation_mutation(
                                "apply theme",
                                "slide layout ids are exhausted".to_owned(),
                            )
                        })?;
                        self.package
                            .content_types
                            .add_override(&part, content_types::SLIDE_LAYOUT);
                        (part, kept)
                    }
                };
                let renamed = copy_relationships(
                    source_package,
                    source_part,
                    &mut self.package,
                    &mut self.media_store,
                    &layout_part,
                    &[rel_types::SLIDE_MASTER],
                    true,
                    &mut layout_relationships,
                )?;
                let xml = rewrite_exact_rel_ids(&source_xml, &renamed)
                    .map_err(|error| malformed(&layout_part, error))?;
                let layout = CT_SlideLayout::from_xml(&xml)
                    .map_err(|error| malformed(&layout_part, error))?;
                self.package.set_part(&layout_part, xml.clone());
                self.package
                    .set_part_rels(&layout_part, layout_relationships);
                match target {
                    Some(index) => {
                        self.layouts[index].layout = layout;
                        self.layouts[index].dirty = true;
                    }
                    None => self.layouts.push(LayoutRecord {
                        part_name: layout_part,
                        layout,
                        dirty: true,
                    }),
                }
            }
            master
                .set_slide_layout_ids(&entries)
                .map_err(|error| malformed(&master_part, error))?;
            self.package.set_part_rels(&master_part, relationships);
            let record = &mut self.masters[master_index];
            record.master = Ok(master);
            record.dirty = true;
        }
        prune_unreachable_parts(&mut self.package, &candidates);
        self.media_store = MediaStore::scan(&self.package);
        Ok(())
    }
}
