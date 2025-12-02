fn main() {
    let icon_path = ".github/logo.ico";

    winresource::WindowsResource::new()
        .set_icon(icon_path)
        .compile()
        .ok();
}
