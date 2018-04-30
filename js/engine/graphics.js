const sleep = async ms => new Promise(resolve => setTimeout(resolve, ms));

const fixImagePath = path =>
    window.location.protocol + "//" + window.location.host + "/assets" + path
        .replace("bmp", "png")
        .replace("gif", "png");

module.exports = () => {
    let _frameTime = 0;

    return {
        Initialize: (hWnd, width, height, colorDepth) => {/*nop*/},

        BeginScene: frameTime => {
            _frameTime = frameTime;
        },

        EndScene: async () => {
            console.log("End scene called.");
            await sleep(_frameTime);
        },

        SetFont: (font, size, bold, italic, underline, strike) => {
            // TODO implement
        },

        DrawText: (text, x, y, color, backcolor) => {
            // TODO implement
        },

        //Public Sub DrawSurface(surfaceID As Long, SrcX As Single, SrcY As Single, SrcWidth As Long, _
        //               SrcHeight As Long, DestX As Single, DestY As Single, _
        //               Optional ByVal DestWidth As Long, Optional ByVal DestHeight As Long, _
        //               Optional ByVal angle As Single = 0, Optional ByVal Alpha As Single = 100, _
        //               Optional ByVal r As Single = 100, Optional ByVal g As Single = 100, _
        //               Optional ByVal b As Single = 100)
        DrawSurface: (id, srcX, srcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight, angle, alpha, r, g, b) => {
            // TODO implement
        },

        CreateSurface: async (path, colorKey, isD3D, notBitmap, reload) => {
            console.log("Graphics::CreateSurface", fixImagePath(path));
            return Math.random();
        },

        DestroySurface: (id) => {/*nop*/}
    };
};