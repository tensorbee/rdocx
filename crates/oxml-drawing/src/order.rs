/// Raw XML children grouped by their schema boundary.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct OrderedRawChildren {
    children: Vec<RawChild>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct RawChild {
    boundary: usize,
    xml: Vec<u8>,
}

impl OrderedRawChildren {
    /// Records a raw child at a caller-defined schema boundary.
    pub fn push(&mut self, boundary: usize, raw_xml: Vec<u8>) {
        self.children.push(RawChild {
            boundary,
            xml: raw_xml,
        });
    }

    /// Returns raw children recorded at a schema boundary.
    pub fn at(&self, boundary: usize) -> impl Iterator<Item = &[u8]> {
        self.children
            .iter()
            .filter(move |child| child.boundary == boundary)
            .map(|child| child.xml.as_slice())
    }

    /// Returns whether no raw children have been recorded.
    pub fn is_empty(&self) -> bool {
        self.children.is_empty()
    }

    /// Keeps only the raw children whose XML satisfies `keep`.
    pub fn retain(&mut self, mut keep: impl FnMut(&[u8]) -> bool) {
        self.children.retain(|child| keep(&child.xml));
    }

    /// Moves raw children at and after one boundary to make room for a new
    /// modelled child without changing their relative order.
    pub fn shift_boundaries_from(&mut self, boundary: usize) {
        for child in &mut self.children {
            if child.boundary >= boundary {
                child.boundary += 1;
            }
        }
    }

    /// Returns raw children at their effective boundary after edits to a
    /// public collection. Each original boundary is anchored to the next
    /// surviving original item, or to the trailing boundary when none remains.
    pub fn at_reconciled<'a>(
        &'a self,
        boundary: usize,
        offset: usize,
        original_to_current: &'a [Option<usize>],
        current_len: usize,
    ) -> impl Iterator<Item = &'a [u8]> {
        self.children
            .iter()
            .filter(move |child| {
                if child.boundary < offset {
                    return false;
                }
                let original_index = child.boundary - offset;
                let effective = original_to_current
                    .iter()
                    .skip(original_index)
                    .flatten()
                    .copied()
                    .next()
                    .unwrap_or(current_len);
                effective == boundary
            })
            .map(|child| child.xml.as_slice())
    }
}

#[cfg(test)]
mod tests {
    use oxml_core::raw_xml::{capture_element, capture_empty_element};
    use quick_xml::events::{BytesEnd, BytesStart, Event};
    use quick_xml::{Reader, Writer};

    use super::OrderedRawChildren;

    struct TestParent {
        has_modelled_child: bool,
        raw_children: OrderedRawChildren,
    }

    impl TestParent {
        fn from_xml(xml: &[u8]) -> Self {
            let mut reader = Reader::from_reader(xml);
            let mut raw_children = OrderedRawChildren::default();
            let mut has_modelled_child = false;
            let mut boundary = 0;
            let mut buffer = Vec::new();

            loop {
                match reader.read_event_into(&mut buffer).unwrap() {
                    Event::Start(element) if local_name(element.name().as_ref()) == b"known" => {
                        has_modelled_child = true;
                        boundary = 1;
                        reader.read_to_end(element.name()).unwrap();
                    }
                    Event::Empty(element) if local_name(element.name().as_ref()) == b"known" => {
                        has_modelled_child = true;
                        boundary = 1;
                    }
                    Event::Start(element) if local_name(element.name().as_ref()) != b"parent" => {
                        let raw = capture_element(&mut reader, &element).unwrap();
                        raw_children.push(boundary, raw);
                    }
                    Event::Empty(element) => {
                        let raw = capture_empty_element(&element).unwrap();
                        raw_children.push(boundary, raw);
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buffer.clear();
            }

            Self {
                has_modelled_child,
                raw_children,
            }
        }

        fn to_xml(&self) -> Vec<u8> {
            let mut writer = Writer::new(Vec::new());
            writer
                .write_event(Event::Start(BytesStart::new("a:parent")))
                .unwrap();
            emit_raw(&mut writer, self.raw_children.at(0));
            if self.has_modelled_child {
                writer
                    .write_event(Event::Empty(BytesStart::new("a:known")))
                    .unwrap();
            }
            emit_raw(&mut writer, self.raw_children.at(1));
            writer
                .write_event(Event::End(BytesEnd::new("a:parent")))
                .unwrap();
            writer.into_inner()
        }
    }

    fn emit_raw<'a>(writer: &mut Writer<Vec<u8>>, children: impl Iterator<Item = &'a [u8]>) {
        for child in children {
            writer.get_mut().extend_from_slice(child);
        }
    }

    fn local_name(name: &[u8]) -> &[u8] {
        name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
    }

    #[test]
    fn modelled_child_between_two_unmodelled_children_keeps_all_three_slots() {
        let parent = TestParent::from_xml(br#"<z:parent><x:first/><z:known/><y:last/></z:parent>"#);

        assert_eq!(
            parent.to_xml(),
            br#"<a:parent><x:first/><a:known/><y:last/></a:parent>"#
        );
    }

    #[test]
    fn multiple_raw_children_at_one_slot_preserve_document_order() {
        let parent =
            TestParent::from_xml(br#"<z:parent><x:first/><y:second/><z:known/></z:parent>"#);

        assert_eq!(
            parent.to_xml(),
            br#"<a:parent><x:first/><y:second/><a:known/></a:parent>"#
        );
    }

    #[test]
    fn raw_subtrees_are_reemitted_byte_for_byte() {
        let parent = TestParent::from_xml(
            br#"<z:parent><x:item x:id="7"><x:child>one &amp; two</x:child><!--note--></x:item><z:known/></z:parent>"#,
        );

        assert_eq!(
            parent.to_xml(),
            br#"<a:parent><x:item x:id="7"><x:child>one &amp; two</x:child><!--note--></x:item><a:known/></a:parent>"#
        );
    }

    #[test]
    fn raw_children_after_the_last_modelled_child_are_not_dropped() {
        let parent =
            TestParent::from_xml(br#"<z:parent><z:known/><x:last x:value="kept"/></z:parent>"#);

        assert!(!parent.raw_children.is_empty());
        assert_eq!(
            parent.to_xml(),
            br#"<a:parent><a:known/><x:last x:value="kept"/></a:parent>"#
        );
    }
}
