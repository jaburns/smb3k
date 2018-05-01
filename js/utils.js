export const pathToURL = path =>
    window.location.href + "assets" + path
        .replace("bmp", "png")
        .replace("gif", "png");

export const sleep = async ms => 
    new Promise(resolve => setTimeout(resolve, ms));