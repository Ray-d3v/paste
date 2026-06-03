use serde::{Deserialize, Serialize};
use std::{
    collections::HashSet,
    env, fs,
    io::{self, Cursor},
    path::{Path, PathBuf},
    time::Duration,
};
use thiserror::Error;
use time::OffsetDateTime;

pub const QUICK_GROUP_ID: &str = "quick";
pub const WORK_GROUP_ID: &str = "work";
pub const IDEA_GROUP_ID: &str = "idea";
pub const MAX_GROUP_NAME_LEN: usize = 32;
pub const DEFAULT_HISTORY_LIMIT: usize = 500;
pub const DEFAULT_TRASH_RETENTION: Duration = Duration::from_secs(60 * 60 * 24 * 7);

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub enum ClipboardKind {
    Text,
    Image,
    Link,
    File,
    Code,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct ClipboardEntry {
    #[serde(default)]
    pub kind: ClipboardKind,
    #[serde(default)]
    pub content: String,
    #[serde(with = "time::serde::rfc3339")]
    pub copied_at: OffsetDateTime,
    #[serde(default)]
    pub source_app: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub image_png_bytes: Option<Vec<u8>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub image_dib_bytes: Option<Vec<u8>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_url: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_title: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub file_paths: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub file_names: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub file_extensions: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_language: Option<String>,
    #[serde(default)]
    pub is_favorite: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pinned_group_id: Option<String>,
    #[serde(default)]
    pub is_deleted: bool,
    #[serde(default, with = "time::serde::rfc3339::option")]
    pub deleted_at: Option<OffsetDateTime>,
}

impl Default for ClipboardKind {
    fn default() -> Self {
        Self::Text
    }
}

impl ClipboardEntry {
    pub fn text(content: impl Into<String>, copied_at: OffsetDateTime) -> Self {
        Self {
            kind: ClipboardKind::Text,
            content: content.into(),
            copied_at,
            source_app: String::new(),
            image_png_bytes: None,
            image_dib_bytes: None,
            link_url: None,
            link_title: None,
            file_paths: Vec::new(),
            file_names: Vec::new(),
            file_extensions: Vec::new(),
            code_language: None,
            is_favorite: false,
            pinned_group_id: None,
            is_deleted: false,
            deleted_at: None,
        }
    }

    pub fn is_pinned(&self) -> bool {
        self.pinned_group_id.is_some()
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct PinnedGroup {
    pub id: String,
    pub name: String,
    pub color_key: String,
    pub sort_order: i32,
}

impl PinnedGroup {
    pub fn new(
        id: impl Into<String>,
        name: impl Into<String>,
        color_key: impl Into<String>,
        sort_order: i32,
    ) -> Self {
        Self {
            id: id.into(),
            name: name.into(),
            color_key: color_key.into(),
            sort_order,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct ClipboardHistoryDocument {
    pub groups: Vec<PinnedGroup>,
    pub entries: Vec<ClipboardEntry>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct RetentionPolicy {
    pub history_limit: usize,
    pub trash_retention: Duration,
}

impl Default for RetentionPolicy {
    fn default() -> Self {
        Self {
            history_limit: DEFAULT_HISTORY_LIMIT,
            trash_retention: DEFAULT_TRASH_RETENTION,
        }
    }
}

#[derive(Debug, Error)]
pub enum HistoryStoreError {
    #[error("local app data directory is unavailable")]
    LocalDataDirectoryUnavailable,
    #[error("history I/O failed")]
    Io(#[from] io::Error),
    #[error("history JSON is invalid")]
    Json(#[from] serde_json::Error),
}

#[derive(Debug, Error)]
pub enum ImageConversionError {
    #[error("DIB data is too short")]
    TooShort,
    #[error("DIB header size {0} is not supported")]
    UnsupportedHeaderSize(u32),
    #[error("DIB dimensions are invalid")]
    InvalidDimensions,
    #[error("DIB compression mode {0} is not supported")]
    UnsupportedCompression(u32),
    #[error("DIB bit depth {0} is not supported")]
    UnsupportedBitDepth(u16),
    #[error("DIB pixel data is truncated")]
    TruncatedPixels,
    #[error("PNG encoding failed")]
    Png(#[from] png::EncodingError),
}

pub fn default_history_path() -> Result<PathBuf, HistoryStoreError> {
    if let Ok(path) = env::var("PASTE_GPUI_HISTORY_PATH") {
        let path = path.trim();
        if !path.is_empty() {
            return Ok(PathBuf::from(path));
        }
    }

    if let Ok(path) = env::var("PASTE_GPUI_DATA_DIR") {
        let path = path.trim();
        if !path.is_empty() {
            return Ok(PathBuf::from(path).join("PasteGPUI").join("history.json"));
        }
    }

    dirs::data_local_dir()
        .map(|path| path.join("PasteGPUI").join("history.json"))
        .ok_or(HistoryStoreError::LocalDataDirectoryUnavailable)
}

pub fn load_history_document(path: &Path) -> Result<ClipboardHistoryDocument, HistoryStoreError> {
    if !path.exists() {
        return Ok(create_document(seeded_groups(), []));
    }

    let text = fs::read_to_string(path)?;
    let document = serde_json::from_str::<ClipboardHistoryDocument>(&text)?;
    Ok(create_document(document.groups, document.entries))
}

pub fn save_history_document(
    path: &Path,
    document: &ClipboardHistoryDocument,
) -> Result<(), HistoryStoreError> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }

    let temp_path = path.with_extension("json.tmp");
    let text = serde_json::to_string_pretty(document)?;
    fs::write(&temp_path, text)?;
    fs::rename(temp_path, path)?;
    Ok(())
}

pub fn seeded_groups() -> Vec<PinnedGroup> {
    vec![
        PinnedGroup::new(QUICK_GROUP_ID, "Quick", "amber", 0),
        PinnedGroup::new(WORK_GROUP_ID, "Work", "sage", 1),
        PinnedGroup::new(IDEA_GROUP_ID, "Idea", "blue", 2),
    ]
}

pub fn create_document(
    groups: impl IntoIterator<Item = PinnedGroup>,
    entries: impl IntoIterator<Item = ClipboardEntry>,
) -> ClipboardHistoryDocument {
    let groups = sanitize_groups(groups);
    let valid_group_ids = groups
        .iter()
        .map(|group| group.id.as_str())
        .collect::<HashSet<_>>();

    let mut entries = entries
        .into_iter()
        .map(|mut entry| {
            let normalized_group_id = normalize_group_id(entry.pinned_group_id.as_deref());
            entry.pinned_group_id =
                normalized_group_id.filter(|id| valid_group_ids.contains(id.as_str()));
            entry
        })
        .collect::<Vec<_>>();

    entries.sort_by(|left, right| right.copied_at.cmp(&left.copied_at));
    ClipboardHistoryDocument { groups, entries }
}

pub fn delete_group(
    groups: impl IntoIterator<Item = PinnedGroup>,
    entries: impl IntoIterator<Item = ClipboardEntry>,
    group_id: &str,
) -> ClipboardHistoryDocument {
    let Some(group_id) = normalize_group_id(Some(group_id)) else {
        return create_document(groups, entries);
    };

    let groups = groups
        .into_iter()
        .filter(|group| group.id != group_id)
        .collect::<Vec<_>>();
    let entries = entries
        .into_iter()
        .map(|mut entry| {
            if entry.pinned_group_id.as_deref() == Some(group_id.as_str()) {
                entry.pinned_group_id = None;
            }
            entry
        })
        .collect::<Vec<_>>();

    create_document(groups, entries)
}

pub fn sanitize_groups(groups: impl IntoIterator<Item = PinnedGroup>) -> Vec<PinnedGroup> {
    let mut groups = groups.into_iter().collect::<Vec<_>>();
    groups.sort_by_key(|group| group.sort_order);

    let mut seen_ids = HashSet::new();
    let mut sanitized = Vec::new();
    for group in groups {
        let Some(id) = normalize_group_id(Some(&group.id)) else {
            continue;
        };
        let name = normalize_group_name(Some(&group.name));
        if name.is_empty() || !seen_ids.insert(id.clone()) {
            continue;
        }

        let sort_order = sanitized.len() as i32;
        let color_key = resolve_color_key(Some(&group.color_key), sanitized.len());
        sanitized.push(PinnedGroup::new(id, name, color_key, sort_order));
    }

    sanitized
}

pub fn create_group(name: &str, groups: &[PinnedGroup]) -> PinnedGroup {
    let existing = sanitize_groups(groups.iter().cloned());
    let id = create_id_from_name(name, &existing);
    let name = normalize_group_name(Some(name));
    let color_key = resolve_color_key(None, existing.len());
    let sort_order = existing
        .iter()
        .map(|group| group.sort_order)
        .max()
        .map_or(0, |value| value + 1);

    PinnedGroup::new(id, name, color_key, sort_order)
}

pub fn rename_group(group: &PinnedGroup, name: &str) -> PinnedGroup {
    PinnedGroup::new(
        group.id.clone(),
        normalize_group_name(Some(name)),
        group.color_key.clone(),
        group.sort_order,
    )
}

pub fn try_validate_group_name(
    name: Option<&str>,
    groups: &[PinnedGroup],
    exclude_id: Option<&str>,
) -> Result<String, String> {
    let name = normalize_group_name(name);
    if name.is_empty() {
        return Err("Name is required.".to_string());
    }

    if name.chars().count() > MAX_GROUP_NAME_LEN {
        return Err(format!(
            "Name must be {MAX_GROUP_NAME_LEN} characters or fewer."
        ));
    }

    let exclude_id = normalize_group_id(exclude_id);
    let exists = groups.iter().any(|group| {
        Some(group.id.as_str()) != exclude_id.as_deref() && group.name.eq_ignore_ascii_case(&name)
    });
    if exists {
        return Err("A group with that name already exists.".to_string());
    }

    Ok(name)
}

pub fn normalize_group_id(group_id: Option<&str>) -> Option<String> {
    let normalized = group_id?.trim().to_lowercase();
    (!normalized.is_empty()
        && normalized
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || ch == '-'))
    .then_some(normalized)
}

pub fn create_id_from_name(name: &str, groups: &[PinnedGroup]) -> String {
    let mut slug = normalize_group_name(Some(name))
        .to_lowercase()
        .chars()
        .map(|ch| if ch.is_ascii_alphanumeric() { ch } else { '-' })
        .collect::<String>();
    while slug.contains("--") {
        slug = slug.replace("--", "-");
    }
    slug = slug.trim_matches('-').to_string();
    if slug.is_empty() {
        slug = "group".to_string();
    }

    let existing_ids = groups
        .iter()
        .map(|group| group.id.as_str())
        .collect::<HashSet<_>>();
    if !existing_ids.contains(slug.as_str()) {
        return slug;
    }

    let mut index = 2;
    loop {
        let candidate = format!("{slug}-{index}");
        if !existing_ids.contains(candidate.as_str()) {
            return candidate;
        }
        index += 1;
    }
}

pub fn is_likely_duplicate(
    existing: &ClipboardEntry,
    incoming: &ClipboardEntry,
    duplicate_merge_window: Duration,
) -> bool {
    let time_diff = if incoming.copied_at >= existing.copied_at {
        incoming.copied_at - existing.copied_at
    } else {
        existing.copied_at - incoming.copied_at
    };
    if time_diff > duplicate_merge_window {
        return false;
    }

    if existing.kind != incoming.kind {
        return false;
    }

    match incoming.kind {
        ClipboardKind::Link => {
            let left = normalize_link_for_comparison(
                existing.link_url.as_deref().unwrap_or(&existing.content),
            );
            let right = normalize_link_for_comparison(
                incoming.link_url.as_deref().unwrap_or(&incoming.content),
            );
            left.eq_ignore_ascii_case(&right)
        }
        ClipboardKind::Image => match (&existing.image_png_bytes, &incoming.image_png_bytes) {
            (Some(left), Some(right)) if !left.is_empty() && !right.is_empty() => left == right,
            _ => match (&existing.image_dib_bytes, &incoming.image_dib_bytes) {
                (Some(left), Some(right)) if !left.is_empty() && !right.is_empty() => left == right,
                _ => false,
            },
        },
        ClipboardKind::File => {
            !existing.file_paths.is_empty()
                && existing.file_paths.len() == incoming.file_paths.len()
                && existing
                    .file_paths
                    .iter()
                    .map(|path| normalize_file_path_for_comparison(path))
                    .eq(incoming
                        .file_paths
                        .iter()
                        .map(|path| normalize_file_path_for_comparison(path)))
        }
        ClipboardKind::Code => existing.content == incoming.content,
        ClipboardKind::Text => existing.content == incoming.content,
    }
}

pub fn merge_entries(existing: &ClipboardEntry, incoming: &ClipboardEntry) -> ClipboardEntry {
    ClipboardEntry {
        kind: incoming.kind.clone(),
        content: choose_preferred_content(&incoming.content, &existing.content),
        copied_at: existing.copied_at.max(incoming.copied_at),
        source_app: if incoming.source_app.trim().is_empty() {
            existing.source_app.clone()
        } else {
            incoming.source_app.clone()
        },
        image_png_bytes: incoming
            .image_png_bytes
            .clone()
            .or_else(|| existing.image_png_bytes.clone()),
        image_dib_bytes: incoming
            .image_dib_bytes
            .clone()
            .or_else(|| existing.image_dib_bytes.clone()),
        link_url: choose_preferred_text(incoming.link_url.as_deref(), existing.link_url.as_deref()),
        link_title: choose_preferred_text(
            incoming.link_title.as_deref(),
            existing.link_title.as_deref(),
        ),
        file_paths: choose_preferred_vec(&incoming.file_paths, &existing.file_paths),
        file_names: choose_preferred_vec(&incoming.file_names, &existing.file_names),
        file_extensions: choose_preferred_vec(&incoming.file_extensions, &existing.file_extensions),
        code_language: choose_preferred_text(
            incoming.code_language.as_deref(),
            existing.code_language.as_deref(),
        ),
        is_favorite: existing.is_favorite || incoming.is_favorite,
        pinned_group_id: existing
            .pinned_group_id
            .clone()
            .or_else(|| incoming.pinned_group_id.clone()),
        is_deleted: existing.is_deleted || incoming.is_deleted,
        deleted_at: existing.deleted_at.or(incoming.deleted_at),
    }
}

pub fn normalize_link_for_comparison(raw_link: &str) -> String {
    let trimmed = raw_link.trim();
    let Some((scheme, rest)) = trimmed.split_once("://") else {
        return trimmed.to_string();
    };

    let (without_fragment, _) = rest.split_once('#').unwrap_or((rest, ""));
    let without_trailing_slash = without_fragment.trim_end_matches('/');
    format!("{scheme}://{without_trailing_slash}")
}

pub fn normalize_file_path_for_comparison(raw_path: &str) -> String {
    raw_path.trim().replace('/', "\\").to_lowercase()
}

pub fn entry_matches_search(entry: &ClipboardEntry, query: &str) -> bool {
    let query = query.trim().to_lowercase();
    if query.is_empty() {
        return true;
    }

    let kind = match entry.kind {
        ClipboardKind::Text => "text",
        ClipboardKind::Image => "image",
        ClipboardKind::Link => "link",
        ClipboardKind::File => "file",
        ClipboardKind::Code => "code",
    };

    contains_casefold(kind, &query)
        || contains_casefold(&entry.source_app, &query)
        || contains_casefold(&entry.content, &query)
        || (entry.is_favorite && contains_casefold("favorite", &query))
        || entry
            .link_url
            .as_deref()
            .is_some_and(|value| contains_casefold(value, &query))
        || entry
            .link_title
            .as_deref()
            .is_some_and(|value| contains_casefold(value, &query))
        || entry
            .file_paths
            .iter()
            .any(|value| contains_casefold(value, &query))
        || entry
            .file_names
            .iter()
            .any(|value| contains_casefold(value, &query))
        || entry
            .file_extensions
            .iter()
            .any(|value| contains_casefold(value, &query))
        || entry
            .code_language
            .as_deref()
            .is_some_and(|value| contains_casefold(value, &query))
}

pub fn visible_entries(
    entries: &[ClipboardEntry],
    is_trash_view: bool,
    active_pinned_group_id: Option<&str>,
    search_query: &str,
) -> Vec<ClipboardEntry> {
    let active_pinned_group_id = normalize_group_id(active_pinned_group_id);
    let mut visible = entries
        .iter()
        .filter(|entry| entry.is_deleted == is_trash_view)
        .filter(|entry| {
            active_pinned_group_id.as_deref().map_or(true, |group_id| {
                entry.pinned_group_id.as_deref() == Some(group_id)
            })
        })
        .filter(|entry| entry_matches_search(entry, search_query))
        .cloned()
        .collect::<Vec<_>>();

    visible.sort_by(|left, right| right.copied_at.cmp(&left.copied_at));

    visible
}

pub fn visible_favorite_entries(
    entries: &[ClipboardEntry],
    is_trash_view: bool,
    search_query: &str,
) -> Vec<ClipboardEntry> {
    let mut visible = entries
        .iter()
        .filter(|entry| entry.is_deleted == is_trash_view)
        .filter(|entry| entry.is_favorite)
        .filter(|entry| entry_matches_search(entry, search_query))
        .cloned()
        .collect::<Vec<_>>();

    visible.sort_by(|left, right| {
        right
            .copied_at
            .cmp(&left.copied_at)
            .then_with(|| right.content.cmp(&left.content))
    });

    visible
}

pub fn visible_kind_entries(
    entries: &[ClipboardEntry],
    is_trash_view: bool,
    kind: ClipboardKind,
    search_query: &str,
) -> Vec<ClipboardEntry> {
    let mut visible = entries
        .iter()
        .filter(|entry| entry.is_deleted == is_trash_view)
        .filter(|entry| entry.kind == kind)
        .filter(|entry| entry_matches_search(entry, search_query))
        .cloned()
        .collect::<Vec<_>>();

    visible.sort_by(|left, right| right.copied_at.cmp(&left.copied_at));

    visible
}

pub fn enforce_retention(
    entries: impl IntoIterator<Item = ClipboardEntry>,
    now: OffsetDateTime,
    policy: RetentionPolicy,
) -> Vec<ClipboardEntry> {
    let mut kept = entries
        .into_iter()
        .filter(|entry| {
            if !entry.is_deleted {
                return true;
            }

            let Some(deleted_at) = entry.deleted_at else {
                return true;
            };
            let Ok(age) = std::time::Duration::try_from(now - deleted_at) else {
                return true;
            };
            age <= policy.trash_retention
        })
        .collect::<Vec<_>>();

    kept.sort_by(|left, right| right.copied_at.cmp(&left.copied_at));

    let mut active_count = 0;
    kept.retain(|entry| {
        if entry.is_deleted {
            return true;
        }

        active_count += 1;
        active_count <= policy.history_limit
    });

    kept
}

pub fn paste_risk_reason(entry: &ClipboardEntry, length_threshold: usize) -> Option<String> {
    let text = match entry.kind {
        ClipboardKind::Text => entry.content.as_str(),
        ClipboardKind::Link => entry.link_url.as_deref().unwrap_or(entry.content.as_str()),
        ClipboardKind::Image => return None,
        ClipboardKind::File => return None,
        ClipboardKind::Code => entry.content.as_str(),
    };

    if text.trim().is_empty() {
        return None;
    }

    if text.len() >= length_threshold {
        return Some(format!("Large paste detected ({} characters).", text.len()));
    }

    if text.contains('\n') {
        return Some("Multi-line content detected.".to_string());
    }

    if looks_like_email(text) || looks_like_token(text) {
        return Some("Sensitive-looking content detected.".to_string());
    }

    None
}

pub fn dib_to_png_bytes(dib: &[u8]) -> Result<Vec<u8>, ImageConversionError> {
    const BITMAPINFOHEADER_SIZE: usize = 40;
    const BI_RGB: u32 = 0;
    const BI_BITFIELDS: u32 = 3;

    if dib.len() < BITMAPINFOHEADER_SIZE {
        return Err(ImageConversionError::TooShort);
    }

    let header_size = read_u32_le(dib, 0)?;
    if header_size != BITMAPINFOHEADER_SIZE as u32 {
        return Err(ImageConversionError::UnsupportedHeaderSize(header_size));
    }

    let width = read_i32_le(dib, 4)?;
    let height = read_i32_le(dib, 8)?;
    let planes = read_u16_le(dib, 12)?;
    let bits_per_pixel = read_u16_le(dib, 14)?;
    let compression = read_u32_le(dib, 16)?;

    if planes != 1 || width <= 0 || height == 0 {
        return Err(ImageConversionError::InvalidDimensions);
    }

    if compression != BI_RGB && !(compression == BI_BITFIELDS && bits_per_pixel == 32) {
        return Err(ImageConversionError::UnsupportedCompression(compression));
    }

    if !matches!(bits_per_pixel, 24 | 32) {
        return Err(ImageConversionError::UnsupportedBitDepth(bits_per_pixel));
    }

    let width = width as usize;
    let height_abs = height.unsigned_abs() as usize;
    let bytes_per_pixel = (bits_per_pixel / 8) as usize;
    let bitfield_masks = if compression == BI_BITFIELDS {
        if dib.len() < BITMAPINFOHEADER_SIZE + 12 {
            return Err(ImageConversionError::TruncatedPixels);
        }
        Some((
            read_u32_le(dib, BITMAPINFOHEADER_SIZE)?,
            read_u32_le(dib, BITMAPINFOHEADER_SIZE + 4)?,
            read_u32_le(dib, BITMAPINFOHEADER_SIZE + 8)?,
        ))
    } else {
        None
    };
    let stride = ((width * bits_per_pixel as usize + 31) / 32) * 4;
    let pixel_offset = header_size as usize + bitfield_masks.map(|_| 12).unwrap_or(0);
    let pixel_bytes = stride
        .checked_mul(height_abs)
        .and_then(|size| pixel_offset.checked_add(size))
        .ok_or(ImageConversionError::InvalidDimensions)?;
    if dib.len() < pixel_bytes {
        return Err(ImageConversionError::TruncatedPixels);
    }

    let top_down = height < 0;
    let mut rgba = Vec::with_capacity(width * height_abs * 4);
    for output_y in 0..height_abs {
        let source_y = if top_down {
            output_y
        } else {
            height_abs - 1 - output_y
        };
        let row_start = pixel_offset + source_y * stride;
        for x in 0..width {
            let pixel_start = row_start + x * bytes_per_pixel;
            let (red, green, blue, alpha) =
                if let Some((red_mask, green_mask, blue_mask)) = bitfield_masks {
                    let pixel =
                        u32::from_le_bytes(dib[pixel_start..pixel_start + 4].try_into().unwrap());
                    (
                        component_from_mask(pixel, red_mask),
                        component_from_mask(pixel, green_mask),
                        component_from_mask(pixel, blue_mask),
                        255,
                    )
                } else {
                    let blue = dib[pixel_start];
                    let green = dib[pixel_start + 1];
                    let red = dib[pixel_start + 2];
                    let alpha = if bits_per_pixel == 32 {
                        dib[pixel_start + 3]
                    } else {
                        255
                    };
                    (red, green, blue, alpha)
                };
            rgba.extend_from_slice(&[red, green, blue, alpha]);
        }
    }

    let mut png_bytes = Vec::new();
    {
        let mut encoder =
            png::Encoder::new(Cursor::new(&mut png_bytes), width as u32, height_abs as u32);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header()?;
        writer.write_image_data(&rgba)?;
    }

    Ok(png_bytes)
}

fn component_from_mask(pixel: u32, mask: u32) -> u8 {
    if mask == 0 {
        return 0;
    }

    let shift = mask.trailing_zeros();
    let max = mask >> shift;
    let value = (pixel & mask) >> shift;
    ((value * 255 + max / 2) / max) as u8
}

fn normalize_group_name(name: Option<&str>) -> String {
    name.unwrap_or_default()
        .trim()
        .chars()
        .take(MAX_GROUP_NAME_LEN)
        .collect::<String>()
        .trim_end()
        .to_string()
}

fn resolve_color_key(color_key: Option<&str>, index: usize) -> String {
    const PALETTE: [&str; 6] = ["amber", "sage", "blue", "slate", "mauve", "olive"];
    let normalized = color_key.unwrap_or_default().trim().to_lowercase();
    if PALETTE.contains(&normalized.as_str()) {
        normalized
    } else {
        PALETTE[index % PALETTE.len()].to_string()
    }
}

fn choose_preferred_content(preferred: &str, fallback: &str) -> String {
    if preferred.trim().is_empty() || matches!(preferred, "[Link]" | "[Image copied]") {
        fallback.to_string()
    } else {
        preferred.to_string()
    }
}

fn choose_preferred_text(preferred: Option<&str>, fallback: Option<&str>) -> Option<String> {
    preferred
        .filter(|value| !value.trim().is_empty())
        .or(fallback)
        .map(ToString::to_string)
}

fn choose_preferred_vec(preferred: &[String], fallback: &[String]) -> Vec<String> {
    if preferred.is_empty() {
        fallback.to_vec()
    } else {
        preferred.to_vec()
    }
}

fn contains_casefold(value: &str, query: &str) -> bool {
    value.to_lowercase().contains(query)
}

fn looks_like_email(text: &str) -> bool {
    text.split_whitespace().any(|part| {
        let Some((local, domain)) = part.split_once('@') else {
            return false;
        };
        !local.is_empty()
            && domain
                .rsplit_once('.')
                .is_some_and(|(_, tld)| tld.len() >= 2)
    })
}

fn looks_like_token(text: &str) -> bool {
    text.split_whitespace().any(|part| {
        part.strip_prefix("sk-")
            .is_some_and(|tail| tail.len() >= 16)
            || part
                .strip_prefix("ghp_")
                .is_some_and(|tail| tail.len() >= 20)
            || part.strip_prefix("AKIA").is_some_and(|tail| {
                tail.len() == 16
                    && tail
                        .chars()
                        .all(|c| c.is_ascii_uppercase() || c.is_ascii_digit())
            })
    })
}

fn read_u16_le(bytes: &[u8], offset: usize) -> Result<u16, ImageConversionError> {
    let value = bytes
        .get(offset..offset + 2)
        .ok_or(ImageConversionError::TooShort)?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32_le(bytes: &[u8], offset: usize) -> Result<u32, ImageConversionError> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or(ImageConversionError::TooShort)?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn read_i32_le(bytes: &[u8], offset: usize) -> Result<i32, ImageConversionError> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or(ImageConversionError::TooShort)?;
    Ok(i32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(seconds: i64) -> OffsetDateTime {
        OffsetDateTime::from_unix_timestamp(seconds).unwrap()
    }

    #[test]
    fn visible_entries_orders_newest_first() {
        let mut old = ClipboardEntry::text("old", at(90));
        old.pinned_group_id = Some("quick".to_string());
        let new = ClipboardEntry::text("new", at(100));

        let visible = visible_entries(&[new.clone(), old.clone()], false, None, "");

        assert_eq!(visible, vec![new, old]);
    }

    #[test]
    fn entry_matches_link_metadata() {
        let mut entry = ClipboardEntry::text("[Link]", at(0));
        entry.kind = ClipboardKind::Link;
        entry.link_title = Some("Rust GPUI".to_string());

        assert!(entry_matches_search(&entry, "gpui"));
    }

    #[test]
    fn duplicate_link_ignores_fragment_and_trailing_slash() {
        let mut existing = ClipboardEntry::text("[Link]", at(0));
        existing.kind = ClipboardKind::Link;
        existing.link_url = Some("https://example.com/path/#section".to_string());
        let mut incoming = ClipboardEntry::text("[Link]", at(5));
        incoming.kind = ClipboardKind::Link;
        incoming.link_url = Some("https://EXAMPLE.com/path".to_string());

        assert!(is_likely_duplicate(
            &existing,
            &incoming,
            Duration::from_secs(10)
        ));
    }

    #[test]
    fn duplicate_file_matches_normalized_paths() {
        let mut existing = ClipboardEntry::text("[Files: 1]", at(0));
        existing.kind = ClipboardKind::File;
        existing.file_paths = vec![r"C:\Temp\Report.pdf".to_string()];
        let mut incoming = ClipboardEntry::text("[Files: 1]", at(1));
        incoming.kind = ClipboardKind::File;
        incoming.file_paths = vec!["c:/temp/report.pdf".to_string()];

        assert!(is_likely_duplicate(
            &existing,
            &incoming,
            Duration::from_secs(10)
        ));
    }

    #[test]
    fn duplicate_code_matches_content() {
        let mut existing = ClipboardEntry::text("fn main() {}", at(0));
        existing.kind = ClipboardKind::Code;
        existing.code_language = Some("rust".to_string());
        let mut incoming = ClipboardEntry::text("fn main() {}", at(1));
        incoming.kind = ClipboardKind::Code;
        incoming.code_language = Some("rust".to_string());

        assert!(is_likely_duplicate(
            &existing,
            &incoming,
            Duration::from_secs(10)
        ));
    }

    #[test]
    fn entry_matches_file_code_and_favorite_metadata() {
        let mut file = ClipboardEntry::text("[Files: 1]", at(0));
        file.kind = ClipboardKind::File;
        file.file_paths = vec![r"C:\Temp\Report.pdf".to_string()];
        file.file_names = vec!["Report.pdf".to_string()];
        file.file_extensions = vec!["pdf".to_string()];

        let mut code = ClipboardEntry::text("Write-Host hi", at(1));
        code.kind = ClipboardKind::Code;
        code.code_language = Some("powershell".to_string());
        code.is_favorite = true;

        assert!(entry_matches_search(&file, "report"));
        assert!(entry_matches_search(&file, "pdf"));
        assert!(entry_matches_search(&code, "powershell"));
        assert!(entry_matches_search(&code, "favorite"));
    }

    #[test]
    fn create_document_unassigns_unknown_group() {
        let groups = vec![PinnedGroup::new("custom", "Custom", "amber", 0)];
        let mut entry = ClipboardEntry::text("kept", at(0));
        entry.pinned_group_id = Some("missing".to_string());

        let document = create_document(groups, [entry]);

        assert_eq!(document.entries[0].pinned_group_id, None);
    }

    #[test]
    fn delete_group_removes_group_and_unassigns_entries() {
        let mut entry = ClipboardEntry::text("kept", at(0));
        entry.pinned_group_id = Some(WORK_GROUP_ID.to_string());

        let document = delete_group(seeded_groups(), [entry], WORK_GROUP_ID);

        assert!(!document
            .groups
            .iter()
            .any(|group| group.id == WORK_GROUP_ID));
        assert_eq!(document.entries[0].pinned_group_id, None);
    }

    #[test]
    fn create_group_slugs_unique_name() {
        let groups = vec![PinnedGroup::new("my-group", "My Group", "amber", 0)];

        let group = create_group("My Group", &groups);

        assert_eq!(group.id, "my-group-2");
    }

    #[test]
    fn rename_group_preserves_identity_and_sanitizes_name() {
        let group = PinnedGroup::new("my-group", "My Group", "amber", 0);

        let renamed = rename_group(&group, "  Updated Group  ");

        assert_eq!(renamed.id, "my-group");
        assert_eq!(renamed.name, "Updated Group");
        assert_eq!(renamed.color_key, "amber");
    }

    #[test]
    fn visible_favorite_entries_filters_favorites() {
        let mut favorite = ClipboardEntry::text("favorite", at(10));
        favorite.is_favorite = true;
        let plain = ClipboardEntry::text("plain", at(20));

        let visible = visible_favorite_entries(&[plain, favorite.clone()], false, "");

        assert_eq!(visible, vec![favorite]);
    }

    #[test]
    fn retention_drops_old_trash_and_caps_active_history() {
        let mut old_trash = ClipboardEntry::text("old trash", at(0));
        old_trash.is_deleted = true;
        old_trash.deleted_at = Some(at(0));
        let active_old = ClipboardEntry::text("active old", at(10));
        let active_new = ClipboardEntry::text("active new", at(20));

        let kept = enforce_retention(
            [old_trash, active_old, active_new.clone()],
            at(60 * 60 * 24 * 8),
            RetentionPolicy {
                history_limit: 1,
                trash_retention: Duration::from_secs(60 * 60 * 24 * 7),
            },
        );

        assert_eq!(kept, vec![active_new]);
    }

    #[test]
    fn paste_risk_detects_multiline_text() {
        let entry = ClipboardEntry::text("one\ntwo", at(0));

        assert_eq!(
            paste_risk_reason(&entry, 10_000),
            Some("Multi-line content detected.".to_string())
        );
    }

    #[test]
    fn history_document_roundtrips_json() {
        let document = create_document(seeded_groups(), [ClipboardEntry::text("hello", at(1))]);
        let json = serde_json::to_string(&document).unwrap();
        let decoded: ClipboardHistoryDocument = serde_json::from_str(&json).unwrap();

        assert_eq!(decoded, document);
    }

    #[test]
    fn legacy_history_json_defaults_extended_fields() {
        let json = r#"{
            "groups": [],
            "entries": [{
                "kind": "Text",
                "content": "hello",
                "copied_at": "1970-01-01T00:00:01Z"
            }]
        }"#;

        let decoded: ClipboardHistoryDocument = serde_json::from_str(json).unwrap();

        assert_eq!(decoded.entries[0].kind, ClipboardKind::Text);
        assert_eq!(decoded.entries[0].source_app, "");
        assert!(decoded.entries[0].file_paths.is_empty());
        assert_eq!(decoded.entries[0].code_language, None);
        assert!(!decoded.entries[0].is_favorite);
    }

    #[test]
    fn history_json_accepts_extended_entries_with_missing_optional_fields() {
        let json = r#"
        {
            "groups": [
                { "id": "quick", "name": "Quick", "color_key": "purple", "sort_order": 0 }
            ],
            "entries": [
                {
                    "kind": "Text",
                    "content": "Meeting notes",
                    "copied_at": "2026-05-31T15:36:20Z",
                    "source_app": "Notes",
                    "is_favorite": false,
                    "is_deleted": false
                },
                {
                    "kind": "Code",
                    "content": "fn main() {}",
                    "copied_at": "2026-05-31T15:36:19Z",
                    "source_app": "Terminal",
                    "code_language": "rust",
                    "is_favorite": true,
                    "is_deleted": false
                },
                {
                    "kind": "File",
                    "content": "C:\\Temp\\proposal.pdf",
                    "copied_at": "2026-05-31T15:36:16Z",
                    "source_app": "Explorer",
                    "file_paths": ["C:\\Temp\\proposal.pdf"],
                    "file_names": ["proposal.pdf"],
                    "file_extensions": ["pdf"],
                    "is_favorite": false,
                    "is_deleted": false
                }
            ]
        }
        "#;

        let decoded: ClipboardHistoryDocument = serde_json::from_str(json).unwrap();

        assert_eq!(decoded.entries.len(), 3);
        assert_eq!(decoded.entries[1].code_language.as_deref(), Some("rust"));
        assert_eq!(decoded.entries[2].file_extensions, vec!["pdf"]);
    }

    #[test]
    fn dib_to_png_converts_bottom_up_24bpp_rgb() {
        let mut dib = bitmap_info_header(2, 2, 24);
        dib.extend_from_slice(&[255, 0, 0, 255, 255, 255, 0, 0]);
        dib.extend_from_slice(&[0, 0, 255, 0, 255, 0, 0, 0]);

        let png = dib_to_png_bytes(&dib).unwrap();

        assert!(png.starts_with(b"\x89PNG\r\n\x1a\n"));
    }

    #[test]
    fn dib_to_png_rejects_unsupported_bit_depth() {
        let dib = bitmap_info_header(1, 1, 8);

        assert!(matches!(
            dib_to_png_bytes(&dib),
            Err(ImageConversionError::UnsupportedBitDepth(8))
        ));
    }

    fn bitmap_info_header(width: i32, height: i32, bits_per_pixel: u16) -> Vec<u8> {
        let mut dib = Vec::new();
        dib.extend_from_slice(&40u32.to_le_bytes());
        dib.extend_from_slice(&width.to_le_bytes());
        dib.extend_from_slice(&height.to_le_bytes());
        dib.extend_from_slice(&1u16.to_le_bytes());
        dib.extend_from_slice(&bits_per_pixel.to_le_bytes());
        dib.extend_from_slice(&0u32.to_le_bytes());
        dib.extend_from_slice(&0u32.to_le_bytes());
        dib.extend_from_slice(&0i32.to_le_bytes());
        dib.extend_from_slice(&0i32.to_le_bytes());
        dib.extend_from_slice(&0u32.to_le_bytes());
        dib.extend_from_slice(&0u32.to_le_bytes());
        dib
    }
}
