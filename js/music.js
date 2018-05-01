import { pathToURL } from './utils';

module.exports = () => {
    let _musicElemsByURL = {};
    let _curElem = null;
    let _playing = false;

    function onMusicEnd() {
        this.currentTime = 0;
        this.play();
    }

    return {
        LoadFile: path => {
            console.log("Music::LoadFile", path);
            const url = pathToURL(path);

            if (! _musicElemsByURL[url]) {
                _musicElemsByURL[url] = new Audio(url);
            }

            _curElem = _musicElemsByURL[url];
        },

        PlayMusic: () => {
            console.log("Music::PlayMusic");
            if (_curElem === null) return;
            if (_playing) return;

            _curElem.addEventListener('ended', onMusicEnd);
            _curElem.play();
            _playing = true;
        },

        StopMusic: () => {
            console.log("Music::StopMusic");
            if (_curElem === null) return;
            if (!_playing) return;

            _curElem.removeEventListener('ended', onMusicEnd);
            _curElem.pause();
            _curElem.currentTime = 0;
            _playing = false;
        }
    };
};