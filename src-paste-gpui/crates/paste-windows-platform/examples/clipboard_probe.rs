use paste_windows_platform::{PlatformIntegration, WindowsPlatformIntegration};

fn main() {
    let platform = WindowsPlatformIntegration::new();
    match platform.read_image_dib_from_clipboard() {
        Ok(Some(bytes)) => {
            println!("image_dib_bytes={}", bytes.len());
            if bytes.len() >= 16 {
                let width = i32::from_le_bytes(bytes[4..8].try_into().unwrap());
                let height = i32::from_le_bytes(bytes[8..12].try_into().unwrap());
                let bit_count = u16::from_le_bytes(bytes[14..16].try_into().unwrap());
                println!("width={width};height={height};bit_count={bit_count}");
            }
        }
        Ok(None) => println!("image_dib_bytes=none"),
        Err(error) => {
            eprintln!("error={error}");
            std::process::exit(1);
        }
    }
}
