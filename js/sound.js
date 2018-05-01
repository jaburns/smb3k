import { pathToURL } from './utils';

module.exports = () => {
    const _soundElems = [];

    return {
        Initialize: hWnd => {
            console.log("Sound::Initialize");
        },

        LoadSound: (path, channel, bits) => {
            console.log("Sound::LoadSound", path);
            const url = pathToURL(path);

            const elem = new Audio(url);
            _soundElems.push(elem);
            return _soundElems.length - 1;
        },

        PlaySound: (id, loop) => {
            _soundElems[id].play();
            _soundElems[id].currentTime = 0;
        },

        StopSound: () => {},

        StillPlaying: id => !_soundElems[id].ended,
    };
};