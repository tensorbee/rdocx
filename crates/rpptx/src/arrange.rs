//! Grouping, ungrouping, aligning, and distributing existing shapes, and
//! gluing connectors to them.
//!
//! Positions follow PowerPoint: a shape flips about the centre of its box,
//! then rotates clockwise about that centre, and a group maps its member
//! space onto its own box before it flips and rotates the same way.

use std::collections::HashSet;
use std::ops::Range;

use oxml_drawing::xfrm::{CT_Point2D, CT_PositiveSize2D, CT_Transform2D};
use rpptx_oxml::connector::{CT_Connection, CT_ConnectionShape};
use rpptx_oxml::shape_tree::{CT_GroupShape, CT_ShapeTree, ShapeIdAllocator, ShapeTreeChild};

use crate::{
    Angle, Emu, Error, Presentation, Result, ShapeKind, ShapeMut, ShapesMut, detach_connectors,
    fit_group_to_members, group_at_mut, invalid_shape_mutation, member_box, shape_kind, shape_mut,
    shape_placeholder, shape_properties, shape_transform,
};

/// The edge or centre line [`ShapesMut::align`] lines shapes up on.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ShapeAlignment {
    Left,
    Center,
    Right,
    Top,
    Middle,
    Bottom,
}

/// What [`ShapesMut::align`] and [`ShapesMut::distribute`] measure against:
/// the bounding box of the shapes given, or the slide.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ArrangeReference {
    Selection,
    Slide,
}

/// The axis [`ShapesMut::distribute`] spaces shapes along.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DistributeDirection {
    Horizontal,
    Vertical,
}

/// One end of a connector, `a:stCxn` for `Begin` and `a:endCxn` for `End`.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ConnectorEnd {
    Begin,
    End,
}

/// An affine map in PDF order: `x' = a x + c y + e`, `y' = b x + d y + f`.
#[derive(Clone, Copy, Debug, PartialEq)]
struct Affine([f64; 6]);

impl Affine {
    const IDENTITY: Self = Self([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]);

    const fn translate(x: f64, y: f64) -> Self {
        Self([1.0, 0.0, 0.0, 1.0, x, y])
    }

    const fn scale(x: f64, y: f64) -> Self {
        Self([x, 0.0, 0.0, y, 0.0, 0.0])
    }

    /// A clockwise rotation on the y-down slide plane.
    fn rotate(degrees: f64) -> Self {
        let (sin, cos) = degrees.to_radians().sin_cos();
        Self([cos, sin, -sin, cos, 0.0, 0.0])
    }

    /// Applies `self`, then `next`.
    fn then(self, next: Self) -> Self {
        let [a, b, c, d, e, f] = self.0;
        let [na, nb, nc, nd, ne, nf] = next.0;
        Self([
            na * a + nc * b,
            nb * a + nd * b,
            na * c + nc * d,
            nb * c + nd * d,
            na * e + nc * f + ne,
            nb * e + nd * f + nf,
        ])
    }

    fn apply(self, (x, y): (f64, f64)) -> (f64, f64) {
        let [a, b, c, d, e, f] = self.0;
        (a * x + c * y + e, b * x + d * y + f)
    }

    fn inverse(self) -> Option<Self> {
        let [a, b, c, d, e, f] = self.0;
        let determinant = a * d - b * c;
        if determinant.abs() < f64::EPSILON {
            return None;
        }
        let (ia, ib, ic, id) = (
            d / determinant,
            -b / determinant,
            -c / determinant,
            a / determinant,
        );
        Some(Self([
            ia,
            ib,
            ic,
            id,
            -(ia * e + ic * f),
            -(ib * e + id * f),
        ]))
    }

    /// Whether the map only translates.
    fn is_translation(self) -> bool {
        let [a, b, c, d, ..] = self.0;
        (a - 1.0).abs() < 1e-9 && b.abs() < 1e-9 && c.abs() < 1e-9 && (d - 1.0).abs() < 1e-9
    }
}

/// Flips about `centre`, then rotates about it, as PowerPoint orients a box.
fn orientation_about(transform: &CT_Transform2D, (x, y): (f64, f64)) -> Affine {
    let flip = |flipped: bool| if flipped { -1.0 } else { 1.0 };
    Affine::translate(-x, -y)
        .then(Affine::scale(
            flip(transform.flip_horizontal),
            flip(transform.flip_vertical),
        ))
        .then(Affine::rotate(transform.rotation.to_degrees()))
        .then(Affine::translate(x, y))
}

fn box_centre(offset: CT_Point2D, extent: CT_PositiveSize2D) -> (f64, f64) {
    (
        offset.x.0 as f64 + extent.cx.0 as f64 / 2.0,
        offset.y.0 as f64 + extent.cy.0 as f64 / 2.0,
    )
}

/// Maps a shape's own box, from `(0, 0)` to its extent, to its parent.
fn box_to_parent(transform: &CT_Transform2D) -> Option<Affine> {
    let (offset, extent) = (transform.offset?, transform.extent?);
    Some(
        Affine::translate(offset.x.0 as f64, offset.y.0 as f64)
            .then(orientation_about(transform, box_centre(offset, extent))),
    )
}

/// The member-space scale of a group on each axis, one where the group has
/// no child extent or a zero one, as the renderer reads it.
fn group_scale(transform: &CT_Transform2D) -> (f64, f64) {
    let ratio = |extent: Option<i64>, child: Option<i64>| match (extent, child) {
        (Some(extent), Some(child)) if child != 0 => extent as f64 / child as f64,
        _ => 1.0,
    };
    (
        ratio(
            transform.extent.map(|extent| extent.cx.0),
            transform.child_extent.map(|extent| extent.cx.0),
        ),
        ratio(
            transform.extent.map(|extent| extent.cy.0),
            transform.child_extent.map(|extent| extent.cy.0),
        ),
    )
}

/// Maps a group's member space to the group's parent.
fn members_to_parent(transform: Option<&CT_Transform2D>) -> Affine {
    let Some(transform) = transform else {
        return Affine::IDENTITY;
    };
    let (Some(offset), Some(extent)) = (transform.offset, transform.extent) else {
        return Affine::IDENTITY;
    };
    let child_offset = transform.child_offset.unwrap_or(offset);
    let (scale_x, scale_y) = group_scale(transform);
    Affine::translate(-(child_offset.x.0 as f64), -(child_offset.y.0 as f64))
        .then(Affine::scale(scale_x, scale_y))
        .then(Affine::translate(offset.x.0 as f64, offset.y.0 as f64))
        .then(orientation_about(transform, box_centre(offset, extent)))
}

/// Maps the member space of the group that `group_path` names, or of the
/// slide for an empty path, to the slide.
fn members_to_slide(children: &[ShapeTreeChild], group_path: &[usize]) -> Option<Affine> {
    let mut maps = Vec::with_capacity(group_path.len());
    let mut level = children;
    for index in group_path {
        let ShapeTreeChild::GroupShape(group) = level.get(*index)? else {
            return None;
        };
        maps.push(members_to_parent(group.group_transform()));
        level = &group.children;
    }
    Some(
        maps.into_iter()
            .rev()
            .fold(Affine::IDENTITY, |inner, outer| inner.then(outer)),
    )
}

/// The box a shape covers in its parent, flips and rotation included, as
/// left, top, right, and bottom.
fn visual_bounds(transform: &CT_Transform2D) -> Option<[f64; 4]> {
    let map = box_to_parent(transform)?;
    let extent = transform.extent?;
    let (width, height) = (extent.cx.0 as f64, extent.cy.0 as f64);
    let corners =
        [(0.0, 0.0), (width, 0.0), (0.0, height), (width, height)].map(|corner| map.apply(corner));
    Some(corners.iter().fold(
        [
            f64::INFINITY,
            f64::INFINITY,
            f64::NEG_INFINITY,
            f64::NEG_INFINITY,
        ],
        |[left, top, right, bottom], (x, y)| {
            [left.min(*x), top.min(*y), right.max(*x), bottom.max(*y)]
        },
    ))
}

fn emu(value: f64) -> Emu {
    Emu(value.round() as i64)
}

/// The members of a slide's own shape tree or of one group on it.
enum Members<'t> {
    Tree(&'t mut CT_ShapeTree),
    Group(&'t mut CT_GroupShape),
}

impl<'t> Members<'t> {
    fn at(tree: &'t mut CT_ShapeTree, group: &[usize]) -> Option<Self> {
        if group.is_empty() {
            Some(Self::Tree(tree))
        } else {
            group_at_mut(&mut tree.children, group).map(Self::Group)
        }
    }

    fn children(&self) -> &[ShapeTreeChild] {
        match self {
            Self::Tree(tree) => &tree.children,
            Self::Group(group) => &group.children,
        }
    }

    fn into_children_mut(self) -> &'t mut Vec<ShapeTreeChild> {
        match self {
            Self::Tree(tree) => &mut tree.children,
            Self::Group(group) => &mut group.children,
        }
    }

    fn children_mut(&mut self) -> &mut Vec<ShapeTreeChild> {
        match self {
            Self::Tree(tree) => &mut tree.children,
            Self::Group(group) => &mut group.children,
        }
    }

    fn take(&mut self, indices: &[usize]) -> Vec<ShapeTreeChild> {
        match self {
            Self::Tree(tree) => tree.take_children(indices),
            Self::Group(group) => group.take_children(indices),
        }
    }

    fn insert(&mut self, index: usize, members: Vec<ShapeTreeChild>) {
        match self {
            Self::Tree(tree) => tree.insert_children(index, members),
            Self::Group(group) => group.insert_children(index, members),
        }
    }
}

/// Returns `indices` ascending, rejecting an empty list, a repeat, and an
/// index outside the `count` members.
fn sorted_indices(operation: &'static str, indices: &[usize], count: usize) -> Result<Vec<usize>> {
    let mut sorted = indices.to_vec();
    sorted.sort_unstable();
    if sorted.is_empty() {
        return Err(invalid_shape_mutation(operation, "no shapes were given"));
    }
    if let Some(index) = sorted.iter().find(|index| **index >= count) {
        return Err(invalid_shape_mutation(
            operation,
            format!("shape index {index} is out of range for {count} shapes"),
        ));
    }
    if let Some(pair) = sorted.windows(2).find(|pair| pair[0] == pair[1]) {
        return Err(invalid_shape_mutation(
            operation,
            format!("shape index {} is given twice", pair[0]),
        ));
    }
    Ok(sorted)
}

impl ShapesMut<'_> {
    fn tree(&mut self) -> &mut CT_ShapeTree {
        &mut self.presentation.slides[self.slide_index]
            .slide
            .common_slide_data
            .shape_tree
    }

    /// Moves existing shapes of this collection into a new group and returns
    /// the group, as python-pptx `add_group_shape(shapes)` does.
    ///
    /// `indices` name z-order positions in this collection, in any order.
    /// The members keep their order among themselves and their place on the
    /// slide: the group's `a:chOff` and `a:chExt` equal its `a:off` and
    /// `a:ext`, the union of the members' boxes. The group takes the z-order
    /// place of the topmost member, so the shapes keep drawing over and
    /// under the same other shapes. Shape ids do not change, so animations
    /// and glued connectors keep their targets. A placeholder, which
    /// PowerPoint does not group, a shape without a position and size of its
    /// own, and a repeated or missing index are rejected without change.
    pub fn group(&mut self, indices: &[usize]) -> Result<ShapeMut<'_>> {
        const OPERATION: &str = "group shapes";
        let group_path = self.group.clone();
        let tree = self.tree();
        let id = ShapeIdAllocator::scan(tree).allocate();
        let mut members = Members::at(tree, &group_path)
            .ok_or_else(|| invalid_shape_mutation(OPERATION, "the collection is not a group"))?;
        let sorted = sorted_indices(OPERATION, indices, members.children().len())?;
        for index in &sorted {
            let child = &members.children()[*index];
            if shape_placeholder(child).is_some() {
                return Err(invalid_shape_mutation(
                    OPERATION,
                    format!(
                        "shape {index} is a placeholder, which PowerPoint does not group; \
                         remove it from the list"
                    ),
                ));
            }
            if member_box(child).is_none() {
                return Err(invalid_shape_mutation(
                    OPERATION,
                    format!("shape {index} has no position and size of its own"),
                ));
            }
        }
        let mut group = CT_GroupShape::new_empty(id, &format!("Group {id}"));
        for member in members.take(&sorted) {
            group.append_child(member);
        }
        fit_group_to_members(&mut group, false);
        let place = sorted[sorted.len() - 1] + 1 - sorted.len();
        members.insert(place, vec![ShapeTreeChild::GroupShape(Box::new(group))]);
        Ok(shape_mut(&mut members.into_children_mut()[place]))
    }

    /// Replaces the group at `index` by its members and returns their new
    /// z-order range.
    ///
    /// Each member keeps its place on the slide: the group's scale, flips,
    /// and rotation move into the member's own transform, as PowerPoint's
    /// Ungroup does. A member that is scaled while rotated by other than a
    /// multiple of 90 degrees keeps its box size along its own axes, which
    /// the file format cannot skew. A table, chart, or other graphic frame
    /// cannot rotate or flip, so a group holding one that is not merely
    /// moved is refused, as is a group that slide animations target.
    /// Connectors glued to the group itself are released, and connectors
    /// glued to members stay glued.
    pub fn ungroup(&mut self, index: usize) -> Result<Range<usize>> {
        const OPERATION: &str = "ungroup shapes";
        let group_path = self.group.clone();
        let slide = &mut self.presentation.slides[self.slide_index].slide;
        let members = Members::at(&mut slide.common_slide_data.shape_tree, &group_path)
            .ok_or_else(|| invalid_shape_mutation(OPERATION, "the collection is not a group"))?;
        let count = members.children().len();
        let child = members.children().get(index).ok_or_else(|| {
            invalid_shape_mutation(
                OPERATION,
                format!("shape index {index} is out of range for {count} shapes"),
            )
        })?;
        let ShapeTreeChild::GroupShape(group) = child else {
            return Err(Error::UnsupportedShapeMutation {
                operation: OPERATION,
                shape_kind: shape_kind(child),
            });
        };
        let group_id = child.non_visual_id();
        if let (Some(timing), Some(id)) = (&slide.timing, group_id)
            && timing.references_shape(id)
        {
            return Err(invalid_shape_mutation(
                OPERATION,
                format!("slide animations target group id {id}; remove them first"),
            ));
        }
        let transform = group.group_transform().cloned().unwrap_or_default();
        let mut placed = group.children.clone();
        for (position, member) in placed.iter_mut().enumerate() {
            place_in_parent(member, &transform).map_err(|message| {
                invalid_shape_mutation(OPERATION, format!("member {position}: {message}"))
            })?;
        }
        let range = index..index + placed.len();
        let mut members = Members::at(&mut slide.common_slide_data.shape_tree, &group_path)
            .expect("the group path was checked");
        members.take(&[index]);
        members.insert(index, placed);
        if let Some(id) = group_id {
            detach_connectors(
                &mut slide.common_slide_data.shape_tree.children,
                &HashSet::from([id]),
            );
        }
        Ok(range)
    }

    /// Lines the shapes at `indices` up on one edge or centre line of their
    /// bounding box, or of the slide, as PowerPoint's Align does.
    ///
    /// Boxes include flips and rotation, so a rotated shape lines up by what
    /// is drawn. A placeholder that inherits its position first receives
    /// that position. `ArrangeReference::Slide` measures in slide
    /// coordinates, so it applies to the slide's own shapes, not to a group's
    /// members.
    pub fn align(
        &mut self,
        indices: &[usize],
        alignment: ShapeAlignment,
        reference: ArrangeReference,
    ) -> Result<()> {
        const OPERATION: &str = "align shapes";
        let (sorted, bounds, target) = self.arrangement(OPERATION, indices, reference)?;
        let shifts = bounds
            .iter()
            .map(|[left, top, right, bottom]| match alignment {
                ShapeAlignment::Left => (target[0] - left, 0.0),
                ShapeAlignment::Center => ((target[0] + target[2] - left - right) / 2.0, 0.0),
                ShapeAlignment::Right => (target[2] - right, 0.0),
                ShapeAlignment::Top => (0.0, target[1] - top),
                ShapeAlignment::Middle => (0.0, (target[1] + target[3] - top - bottom) / 2.0),
                ShapeAlignment::Bottom => (0.0, target[3] - bottom),
            })
            .collect::<Vec<_>>();
        self.shift(OPERATION, &sorted, &shifts)
    }

    /// Spaces the shapes at `indices` evenly along one axis, as PowerPoint's
    /// Distribute does.
    ///
    /// Shapes are taken in their order along the axis. Relative to the
    /// selection, the first and last stay and the others move so that the
    /// gaps between boxes are equal, which needs three shapes. Relative to
    /// the slide, the first box starts at the slide edge and the last ends
    /// at the other, and a single shape is centred. Boxes include flips and
    /// rotation, as for [`Self::align`].
    pub fn distribute(
        &mut self,
        indices: &[usize],
        direction: DistributeDirection,
        reference: ArrangeReference,
    ) -> Result<()> {
        const OPERATION: &str = "distribute shapes";
        let (sorted, bounds, target) = self.arrangement(OPERATION, indices, reference)?;
        let (start, end) = match direction {
            DistributeDirection::Horizontal => (0, 2),
            DistributeDirection::Vertical => (1, 3),
        };
        let mut order = (0..sorted.len()).collect::<Vec<_>>();
        order.sort_by(|left, right| bounds[*left][start].total_cmp(&bounds[*right][start]));
        let lengths = bounds
            .iter()
            .map(|bound| bound[end] - bound[start])
            .collect::<Vec<_>>();
        let total = lengths.iter().sum::<f64>();
        let span = target[end] - target[start];
        let mut shifts = vec![(0.0, 0.0); sorted.len()];
        let mut cursor = target[start];
        let gap = if order.len() == 1 {
            cursor += (span - total) / 2.0;
            0.0
        } else {
            (span - total) / (order.len() - 1) as f64
        };
        for position in order {
            let delta = cursor - bounds[position][start];
            shifts[position] = match direction {
                DistributeDirection::Horizontal => (delta, 0.0),
                DistributeDirection::Vertical => (0.0, delta),
            };
            cursor += lengths[position] + gap;
        }
        self.shift(OPERATION, &sorted, &shifts)
    }

    /// Checks the shapes an arrangement moves and returns their sorted
    /// indices, their drawn boxes, and the box they arrange against.
    #[allow(clippy::type_complexity)]
    fn arrangement(
        &mut self,
        operation: &'static str,
        indices: &[usize],
        reference: ArrangeReference,
    ) -> Result<(Vec<usize>, Vec<[f64; 4]>, [f64; 4])> {
        if reference == ArrangeReference::Slide && !self.group.is_empty() {
            return Err(invalid_shape_mutation(
                operation,
                "a group's members arrange relative to the selection, not the slide",
            ));
        }
        let group_path = self.group.clone();
        let count =
            Members::at(self.tree(), &group_path).map_or(0, |members| members.children().len());
        let sorted = sorted_indices(operation, indices, count)?;
        #[cfg(feature = "render")]
        for index in &sorted {
            let mut path = self.group.clone();
            path.push(*index);
            self.presentation
                .materialize_geometry(self.slide_index, &path)?;
        }
        let members = Members::at(self.tree(), &group_path).expect("the group path was checked");
        let bounds = sorted
            .iter()
            .map(|index| {
                let child = &members.children()[*index];
                if matches!(shape_kind(child), ShapeKind::AlternateContent) {
                    return Err(Error::UnsupportedShapeMutation {
                        operation,
                        shape_kind: ShapeKind::AlternateContent,
                    });
                }
                shape_transform(child)
                    .and_then(visual_bounds)
                    .ok_or_else(|| {
                        invalid_shape_mutation(
                            operation,
                            format!("shape {index} has no position and size"),
                        )
                    })
            })
            .collect::<Result<Vec<_>>>()?;
        let target = match reference {
            ArrangeReference::Selection => bounds.iter().fold(
                [
                    f64::INFINITY,
                    f64::INFINITY,
                    f64::NEG_INFINITY,
                    f64::NEG_INFINITY,
                ],
                |[left, top, right, bottom], bound| {
                    [
                        left.min(bound[0]),
                        top.min(bound[1]),
                        right.max(bound[2]),
                        bottom.max(bound[3]),
                    ]
                },
            ),
            ArrangeReference::Slide => {
                let (width, height) = self.presentation.slide_size().ok_or_else(|| {
                    invalid_shape_mutation(operation, "the presentation has no slide size")
                })?;
                [0.0, 0.0, width.0 as f64, height.0 as f64]
            }
        };
        Ok((sorted, bounds, target))
    }

    /// Moves each shape at `sorted` by its `(dx, dy)`.
    fn shift(
        &mut self,
        operation: &'static str,
        sorted: &[usize],
        shifts: &[(f64, f64)],
    ) -> Result<()> {
        let group_path = self.group.clone();
        let mut members =
            Members::at(self.tree(), &group_path).expect("the group path was checked");
        for (index, (dx, dy)) in sorted.iter().zip(shifts) {
            let mut shape = shape_mut(&mut members.children_mut()[*index]);
            let transform = shape.transform_mut(operation)?;
            let offset = transform.offset.unwrap_or_default();
            transform.offset = Some(CT_Point2D {
                x: Emu(offset.x.0 + dx.round() as i64),
                y: Emu(offset.y.0 + dy.round() as i64),
            });
        }
        Ok(())
    }
}

/// Rewrites one member's transform so that it draws in its group's parent
/// where it drew inside the group of `transform`.
fn place_in_parent(
    member: &mut ShapeTreeChild,
    transform: &CT_Transform2D,
) -> std::result::Result<(), String> {
    let map = members_to_parent(Some(transform));
    if map == Affine::IDENTITY {
        return Ok(());
    }
    let kind = shape_kind(member);
    if kind == ShapeKind::GraphicFrame && !map.is_translation() {
        return Err(
            "a table, chart, or other graphic frame cannot take the group's scale, flips, or \
             rotation; reset them on the group first"
                .to_owned(),
        );
    }
    let own = shape_transform(member)
        .cloned()
        .ok_or_else(|| "the member has no position and size of its own".to_owned())?;
    let (Some(offset), Some(extent)) = (own.offset, own.extent) else {
        return Err("the member has no position and size of its own".to_owned());
    };
    let (centre_x, centre_y) = map.apply(box_centre(offset, extent));
    let (scale_x, scale_y) = group_scale(transform);
    let own_degrees = own.rotation.to_degrees().rem_euclid(180.0);
    let (scale_width, scale_height) = if (45.0..135.0).contains(&own_degrees) {
        (scale_y, scale_x)
    } else {
        (scale_x, scale_y)
    };
    let width = extent.cx.0 as f64 * scale_width;
    let height = extent.cy.0 as f64 * scale_height;
    let mirrored = transform.flip_horizontal != transform.flip_vertical;
    let rotation = if mirrored {
        i64::from(transform.rotation.0) - i64::from(own.rotation.0)
    } else {
        i64::from(transform.rotation.0) + i64::from(own.rotation.0)
    }
    .rem_euclid(21_600_000);
    let mut shape = shape_mut(member);
    let placed = shape
        .transform_mut("ungroup shapes")
        .map_err(|error| error.to_string())?;
    placed.offset = Some(CT_Point2D {
        x: emu(centre_x - width / 2.0),
        y: emu(centre_y - height / 2.0),
    });
    placed.extent = Some(CT_PositiveSize2D {
        cx: emu(width),
        cy: emu(height),
    });
    if kind != ShapeKind::GraphicFrame {
        placed.rotation = Angle(i32::try_from(rotation).expect("a normalized angle fits in i32"));
        placed.flip_horizontal = own.flip_horizontal != transform.flip_horizontal;
        placed.flip_vertical = own.flip_vertical != transform.flip_vertical;
    }
    Ok(())
}

/// Returns the z-order path of the only shape on a slide with `shape_id`.
fn path_of_id(children: &[ShapeTreeChild], shape_id: u32) -> Option<Vec<usize>> {
    for (index, child) in children.iter().enumerate() {
        if child.non_visual_id() == Some(shape_id) {
            return Some(vec![index]);
        }
        if let ShapeTreeChild::GroupShape(group) = child
            && let Some(mut path) = path_of_id(&group.children, shape_id)
        {
            path.insert(0, index);
            return Some(path);
        }
    }
    None
}

fn child_at<'t>(children: &'t [ShapeTreeChild], path: &[usize]) -> Option<&'t ShapeTreeChild> {
    let (last, groups) = path.split_last()?;
    let mut level = children;
    for index in groups {
        let ShapeTreeChild::GroupShape(group) = level.get(*index)? else {
            return None;
        };
        level = &group.children;
    }
    level.get(*last)
}

fn child_at_mut<'t>(
    children: &'t mut [ShapeTreeChild],
    path: &[usize],
) -> Option<&'t mut ShapeTreeChild> {
    let (last, groups) = path.split_last()?;
    if groups.is_empty() {
        return children.get_mut(*last);
    }
    group_at_mut(children, groups)?.children.get_mut(*last)
}

/// Rewrites a connector's transform so that it runs from `begin` to
/// `finish`, in its parent's coordinates, without rotation.
fn place_connector_ends(connector: &mut CT_ConnectionShape, begin: (f64, f64), finish: (f64, f64)) {
    let transform = connector
        .shape_properties
        .transform
        .get_or_insert_with(CT_Transform2D::default);
    transform.offset = Some(CT_Point2D {
        x: emu(begin.0.min(finish.0)),
        y: emu(begin.1.min(finish.1)),
    });
    transform.extent = Some(CT_PositiveSize2D {
        cx: emu((begin.0 - finish.0).abs()),
        cy: emu((begin.1 - finish.1).abs()),
    });
    transform.rotation = Angle(0);
    transform.flip_horizontal = begin.0 > finish.0;
    transform.flip_vertical = begin.1 > finish.1;
}

impl ShapeMut<'_> {
    /// Moves one end of a connector to `(x, y)` in its parent's
    /// coordinates, as python-pptx `begin_x`, `begin_y`, `end_x`, and
    /// `end_y` assignments do, and releases that end's glue, so PowerPoint
    /// does not route it back to the shape it was glued to. The other end
    /// stays where it is, and the connector's rotation resets to none.
    /// Other shapes are rejected.
    pub fn set_connector_endpoint(&mut self, end: ConnectorEnd, x: Emu, y: Emu) -> Result<()> {
        const OPERATION: &str = "set connector endpoint";
        let ShapeTreeChild::Connector(connector) = &mut *self.child else {
            return Err(Error::UnsupportedShapeMutation {
                operation: OPERATION,
                shape_kind: shape_kind(self.child),
            });
        };
        let point = (x.0 as f64, y.0 as f64);
        let current = connector
            .shape_properties
            .transform
            .as_ref()
            .and_then(connector_endpoints)
            .unwrap_or([point, point]);
        match end {
            ConnectorEnd::Begin => {
                place_connector_ends(connector, point, current[1]);
                connector.start_connection = None;
            }
            ConnectorEnd::End => {
                place_connector_ends(connector, current[0], point);
                connector.end_connection = None;
            }
        }
        Ok(())
    }
}

/// The begin and end points of a connector in its parent's coordinates.
pub(crate) fn connector_endpoints(transform: &CT_Transform2D) -> Option<[(f64, f64); 2]> {
    let map = box_to_parent(transform)?;
    let extent = transform.extent?;
    Some([
        map.apply((0.0, 0.0)),
        map.apply((extent.cx.0 as f64, extent.cy.0 as f64)),
    ])
}

impl Presentation {
    /// Glues one end of the connector at `connector_path` to connection
    /// site `site` of the shape whose `p:cNvPr/@id` is `shape_id`, as
    /// python-pptx `begin_connect` and `end_connect` do.
    ///
    /// `connector_path` holds one z-order index per nesting level, as for
    /// [`Self::effective_geometry`]. The end is recorded in `a:stCxn` or
    /// `a:endCxn` and moves to the site: the preset geometry's own
    /// connection sites with its adjustments, flips, and rotation, or the
    /// four edge midpoints, top, left, bottom, and right, of a shape without
    /// a preset or whose preset defines none, as python-pptx counts them. The other end stays where it
    /// is, and the connector's rotation resets to none. A site outside the
    /// shape's sites, a missing or repeated shape id, and a path that names
    /// no connector are rejected without change.
    pub fn connect_connector(
        &mut self,
        slide_index: usize,
        connector_path: &[usize],
        end: ConnectorEnd,
        shape_id: u32,
        site: u32,
    ) -> Result<()> {
        const OPERATION: &str = "connect connector";
        self.require_slide_index(slide_index)?;
        let children = &self.slides[slide_index]
            .slide
            .common_slide_data
            .shape_tree
            .children;
        let connector = child_at(children, connector_path).ok_or_else(|| {
            invalid_shape_mutation(
                OPERATION,
                format!("slide {slide_index} has no shape at path {connector_path:?}"),
            )
        })?;
        let ShapeTreeChild::Connector(connector) = connector else {
            return Err(Error::UnsupportedShapeMutation {
                operation: OPERATION,
                shape_kind: shape_kind(connector),
            });
        };
        if crate::shape_id_count(children, shape_id) != 1 {
            return Err(invalid_shape_mutation(
                OPERATION,
                format!("shape id {shape_id} is missing or not unique on the slide"),
            ));
        }
        let target_path = path_of_id(children, shape_id).ok_or_else(|| {
            invalid_shape_mutation(
                OPERATION,
                format!("shape id {shape_id} is inside an mc:AlternateContent fallback"),
            )
        })?;
        if target_path == connector_path {
            return Err(invalid_shape_mutation(
                OPERATION,
                "a connector cannot connect to itself",
            ));
        }
        let target = child_at(children, &target_path).expect("the path was found");
        let target_transform = match shape_transform(target) {
            Some(transform) if transform.offset.is_some() && transform.extent.is_some() => {
                transform.clone()
            }
            _ => self.inherited_target_transform(slide_index, target, OPERATION)?,
        };
        let extent = target_transform
            .extent
            .expect("the transform has an extent");
        let size = (extent.cx.0 as f64, extent.cy.0 as f64);
        let sites = shape_properties(target)
            .and_then(|properties| properties.preset_geometry.as_ref())
            .map(|geometry| geometry.connection_sites(size))
            .transpose()
            .map_err(|error| invalid_shape_mutation(OPERATION, error.to_string()))?
            .flatten()
            .filter(|sites| !sites.is_empty())
            .unwrap_or_else(|| {
                vec![
                    (size.0 / 2.0, 0.0),
                    (0.0, size.1 / 2.0),
                    (size.0 / 2.0, size.1),
                    (size.0, size.1 / 2.0),
                ]
            });
        let local = *sites.get(site as usize).ok_or_else(|| {
            invalid_shape_mutation(
                OPERATION,
                format!(
                    "shape id {shape_id} has {} connection sites, 0 to {}, not {site}",
                    sites.len(),
                    sites.len().saturating_sub(1)
                ),
            )
        })?;
        let on_slide = members_to_slide(children, &target_path[..target_path.len() - 1])
            .zip(box_to_parent(&target_transform))
            .map(|(parent, own)| own.then(parent).apply(local))
            .ok_or_else(|| invalid_shape_mutation(OPERATION, "the shape has no position"))?;
        let point = members_to_slide(children, &connector_path[..connector_path.len() - 1])
            .and_then(Affine::inverse)
            .map(|map| map.apply(on_slide))
            .ok_or_else(|| {
                invalid_shape_mutation(
                    OPERATION,
                    "the connector's group has no size, so no point maps into it",
                )
            })?;
        let current = connector
            .shape_properties
            .transform
            .as_ref()
            .and_then(connector_endpoints)
            .unwrap_or([point, point]);
        let [begin, finish] = match end {
            ConnectorEnd::Begin => [point, current[1]],
            ConnectorEnd::End => [current[0], point],
        };

        let children = &mut self.slides[slide_index]
            .slide
            .common_slide_data
            .shape_tree
            .children;
        let Some(ShapeTreeChild::Connector(connector)) = child_at_mut(children, connector_path)
        else {
            unreachable!("the connector was found above");
        };
        place_connector_ends(connector, begin, finish);
        let connection = Some(CT_Connection::new(shape_id, site));
        match end {
            ConnectorEnd::Begin => connector.start_connection = connection,
            ConnectorEnd::End => connector.end_connection = connection,
        }
        Ok(())
    }

    #[cfg(feature = "render")]
    fn inherited_target_transform(
        &self,
        slide_index: usize,
        target: &ShapeTreeChild,
        operation: &'static str,
    ) -> Result<CT_Transform2D> {
        let inherited = match shape_placeholder(target) {
            Some(placeholder) => self.inherited_transform(slide_index, placeholder)?,
            None => None,
        };
        inherited
            .filter(|transform| transform.offset.is_some() && transform.extent.is_some())
            .ok_or_else(|| invalid_shape_mutation(operation, "the shape has no position and size"))
    }

    #[cfg(not(feature = "render"))]
    fn inherited_target_transform(
        &self,
        _slide_index: usize,
        _target: &ShapeTreeChild,
        operation: &'static str,
    ) -> Result<CT_Transform2D> {
        Err(invalid_shape_mutation(
            operation,
            "the shape has no position and size of its own",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::{Affine, members_to_parent, place_in_parent};
    use oxml_core::units::{Angle, Emu};
    use oxml_drawing::xfrm::{CT_Point2D, CT_PositiveSize2D, CT_Transform2D};

    fn transform(x: i64, y: i64, cx: i64, cy: i64) -> CT_Transform2D {
        let mut transform = CT_Transform2D::default();
        transform.offset = Some(CT_Point2D {
            x: Emu(x),
            y: Emu(y),
        });
        transform.extent = Some(CT_PositiveSize2D {
            cx: Emu(cx),
            cy: Emu(cy),
        });
        transform
    }

    #[test]
    fn affine_inverse_undoes_a_rotated_scaled_map() {
        let mut group = transform(100, 200, 400, 300);
        group.child_offset = Some(CT_Point2D {
            x: Emu(0),
            y: Emu(0),
        });
        group.child_extent = Some(CT_PositiveSize2D {
            cx: Emu(200),
            cy: Emu(100),
        });
        group.rotation = Angle(30 * 60_000);
        group.flip_horizontal = true;
        let map = members_to_parent(Some(&group));
        let back = map.inverse().unwrap().apply(map.apply((37.0, 81.0)));
        assert!((back.0 - 37.0).abs() < 1e-6 && (back.1 - 81.0).abs() < 1e-6);
        assert_eq!(members_to_parent(None), Affine::IDENTITY);
    }

    #[test]
    fn a_member_of_a_rotated_mirrored_group_keeps_its_drawn_box() {
        let mut group = transform(0, 0, 200, 100);
        group.child_offset = Some(CT_Point2D {
            x: Emu(0),
            y: Emu(0),
        });
        group.child_extent = Some(CT_PositiveSize2D {
            cx: Emu(200),
            cy: Emu(100),
        });
        group.rotation = Angle(90 * 60_000);
        group.flip_horizontal = true;
        let mut member = rpptx_oxml::shape_tree::ShapeTreeChild::Shape(
            rpptx_oxml::shape_tree::CT_Shape::new_preset(2, "Shape 2", "rect", {
                let mut own = transform(0, 0, 40, 20);
                own.rotation = Angle(10 * 60_000);
                own
            })
            .unwrap(),
        );
        place_in_parent(&mut member, &group).unwrap();
        let rpptx_oxml::shape_tree::ShapeTreeChild::Shape(shape) = member else {
            unreachable!();
        };
        let placed = shape.shape_properties.transform.unwrap();
        // The member's centre (20, 10) mirrors to (180, 10), then turns a
        // quarter about the group centre (100, 50) to (140, 130).
        assert_eq!(
            placed.offset,
            Some(CT_Point2D {
                x: Emu(120),
                y: Emu(120)
            })
        );
        assert_eq!(
            placed.extent,
            Some(CT_PositiveSize2D {
                cx: Emu(40),
                cy: Emu(20)
            })
        );
        assert_eq!(placed.rotation, Angle(80 * 60_000));
        assert!(placed.flip_horizontal && !placed.flip_vertical);
    }
}
