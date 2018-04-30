
export const pathToURL = path =>
    window.location.protocol + "//" + window.location.host + "/assets" + path
        .replace("bmp", "png")
        .replace("gif", "png");

export const sleep = async ms => 
    new Promise(resolve => setTimeout(resolve, ms));