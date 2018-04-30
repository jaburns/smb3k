import { pathToURL, sleep } from '../utils';

//const log = console.log.bind(console);
const log = () => {};

module.exports = () => {
    let _frameTime = 0;
    let _loadedSurfaces = [];
    let _ctx = document.querySelector('#game').getContext('2d');

    const loadSurfaceImage = async url => new Promise(resolve => {
        let img = new Image();
        img.onload = () => {
            _loadedSurfaces.push(img);
            resolve(_loadedSurfaces.length - 1);
        };
        img.src = url;
    });

    return {
        Initialize: (hWnd, width, height, colorDepth) => {
            log("Graphics::Initialize", width, height, colorDepth);
        },

        BeginScene: frameTime => {
            log("Graphics::BeginScene", frameTime);
            _frameTime = frameTime;
        },

        EndScene: async () => {
            log("Graphics::EndScene");
            await sleep(_frameTime);
        },

        SetFont: (font, size, bold, italic, underline, strike) => {
            log("Graphics::SetFont", font, size, bold, italic, underline, strike);

            _ctx.font = size + "pt " + font;
        },

        DrawText: (text, x, y, color, backcolor) => {
            log("Graphics::DrawText", text, x, y, color, backcolor);

            if (typeof color !== 'undefined') _ctx.fillStyle = color;

            _ctx.fillText(text, x + 5, y + 10);
        },

        DrawRect: (x, y, width, height, color) => {
            log("Graphics::DrawRect", x, y, width, height, color);

            _ctx.fillStyle = '#000';
            _ctx.fillRect(x, y, width, height);
        },

        DrawSurface: (id, srcX, srcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight, angle, alpha, r, g, b) => {
            log("Graphics::DrawSurface", id, srcX, srcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight, angle, alpha, r, g, b);

            if (typeof destWidth === 'undefined') destWidth = srcWidth;
            if (typeof destHeight === 'undefined') destHeight = srcHeight;
            if (typeof alpha !== 'undefined') _ctx.globalAlpha = alpha / 100;

            if (typeof angle !== 'undefined') {
                _ctx.save();
                _ctx.translate(destX + destWidth / 2, destY + destHeight / 2);
                _ctx.rotate(angle * Math.PI / 180);
                _ctx.drawImage(_loadedSurfaces[id], srcX, srcY, srcWidth, srcHeight, -destWidth / 2, -destHeight / 2, destWidth, destHeight);
                _ctx.restore();
            }
            else {
                _ctx.drawImage(_loadedSurfaces[id], srcX, srcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight);
            }

            if (typeof alpha !== 'undefined') _ctx.globalAlpha = 1;
        },

        CreateSurface: async (path, colorKey, isD3D, notBitmap, reload) => {
            let url = pathToURL(path);
            log("Graphics::CreateSurface", pathToURL(path));
            let id = await loadSurfaceImage(url);
            log("Graphics::CreateSurface", "Created surface with id:", id);
            return id;
        },

        DestroySurface: (id) => {
            log("Graphics::DestroySurface", id);
        }
    };
};