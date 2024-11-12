var fast =  require('C:\\Program Files (x86)\\domotz\\lib\\node_modules\\domotz-remote-pawn-ng\\src\\network_tools\\fast');

function test() {
    var onProgress = function (progress) {
        if (progress.testType === 'latency') {
            // latency test reports result in ms in the 'value' field
            console.debug('Progress: ' + progress.value + 'ms');
        } else {
            // download/upload tests report result in the 'speed' field
            var mbs = progress.speed / 1000 / 1000;
            console.debug('Progress: ' + mbs + 'Mbps');
        }
    };

    var results = {};

    var onSnapshot = function (result) {
        var speed = result.bytes / result.time * 1000 * 8,
            mbs = speed / 1000 / 1000;
        console.debug('Snapshot speed: ' + mbs + 'Mbps');
    };

    var testerConfig = {
            duration: { // the test duration will be between min/max based on stability of measurements
                min: 5, // minimum test duration
                max: 30, // maximum test duration
            },
            connections: {
                min: 1,  // minimum number of parallel connections
                max: 8,  // maximum number of parallel connections
            },
            https: false, // whether to run the test using https
        },
        testerContext = {
            name: 'Node.js Domotz Agent', // name of the testing app. Used for logging
            deviceType: "Windows", // type of the device running the test
            version:"test"}

        tester = fast.tester(testerConfig, testerContext),
        events = fast.event;

    var onCompleteLatency = function (result) {
        console.debug('Final latency: ' + result.value + 'ms');
        results.latency = result.value;
        console.debug(JSON.stringify(results, null, 2));
        process.exit()
    };

    var onCompleteUpload = function (result) {
        results.upload = result.speed;
        tester.off(events.END, onCompleteUpload);
        tester.on(events.END, onCompleteLatency);
        console.debug('Latency test start');
        results.latency = 0;
        tester.latency();
    };

    var onCompleteDownload = function (result) {
        results.download = result.speed;
        tester.off(events.END, onCompleteDownload);
        tester.on(events.END, onCompleteUpload);
        console.debug('Upload test start');
        tester.upload();
    };

    var onFail = function (error) {
        console.error("ERROR:" + error);
    };

    // set up events
    // events.START - indicates start of the test
    // events.END - indicates the end of the test
    // events.PROGRESS - shows intermediate aggregate speed measurement (since start)
    // events.SNAPSHOT - shows instant speed measurement (since last snapshot)
    tester.on(events.PROGRESS, onProgress);

    tester.on(events.SNAPSHOT, onSnapshot);

    tester.on(events.END, onCompleteDownload);

    tester.on(events.FAIL, onFail);

    tester.on(events.CONNECTION_FAIL, onFail);

    console.debug('Starting FAST speed test, download');

    tester.download();
}

test()