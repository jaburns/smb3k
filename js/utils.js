export const pathToURL = path =>
    window.location.href + "assets" + path
        .replace(".bmp", ".png")
        .replace(".gif", ".png")
        .replace(".mid", ".mp3");

export const sleep = async ms => 
    new Promise(resolve => setTimeout(resolve, ms));