use paste_core::dib_to_png_bytes;
use paste_windows_platform::{PlatformIntegration, WindowsPlatformIntegration};

fn main() {
    let platform = WindowsPlatformIntegration::new();
    match platform.read_image_dib_from_clipboard() {
        Ok(Some(bytes)) => {
            println!("image_dib_bytes={}", bytes.len());
            match dib_to_png_bytes(&bytes) {
                Ok(png) => println!("image_png_bytes={}", png.len()),
                Err(error) => {
                    eprintln!("png_error={error}");
                    std::process::exit(2);
                }
            }
        }
        Ok(None) => println!("image_dib_bytes=none"),
        Err(error) => {
            eprintln!("dib_error={error}");
            std::process::exit(1);
        }
    }
}
