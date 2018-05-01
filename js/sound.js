import { pathToURL } from './utils';

export const loadSoundElem = (url, collection) => {
    const elem = new Audio(url);
    collection.push(elem);
    return collection.length - 1;
};

module.exports = () => {
    const _soundElems = [];

    return {
        Initialize: hWnd => {
            console.log("Sound::Initialize");
        },

        LoadSound: (path, channel, bits) => {
            console.log("Sound::LoadSound", path);
            const url = pathToURL(path);
            return loadSoundElem(url, _soundElems);
        },

        PlaySound: (id, loop) => {
            _soundElems[id].play();
            _soundElems[id].currentTime = 0;
        },

        StopSound: () => {},

        StillPlaying: id => !_soundElems[id].ended,
    };
};