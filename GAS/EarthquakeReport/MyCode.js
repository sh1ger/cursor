// 定数定義
const CONSTANTS = {
  FEED_URL: 'https://www.data.jma.go.jp/developer/xml/feed/eqvol.xml',
  NAMESPACES: {
    ATOM: 'http://www.w3.org/2005/Atom',
    SEISMOLOGY: 'http://xml.kishou.go.jp/jmaxml1/'
  },
  CONTENT_TYPE: '震度情報'
};

/**
 * 気象庁の地震情報フィードからデータを取得し、スプレッドシートに記録する
 * @returns {Array<Array<string>>|null} 地震情報の配列、またはデータが存在しない場合はnull
 */
function fetchEarthquakeData() {
  try {
    const feed = fetchAndParseFeed();
    const earthQuakeContents = filterEarthquakeEntries(feed);
    
    if (!earthQuakeContents?.length) {
      Logger.log('地震情報が見つかりませんでした');
      return null;
    }

    const targetEarthQuakes = processEarthquakeContents(earthQuakeContents);
    saveToSpreadsheet(targetEarthQuakes);
    
    return targetEarthQuakes;
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    throw error;
  }
}

/**
 * フィードを取得してパースする
 * @returns {XMLDocument} パースされたXMLドキュメント
 */
function fetchAndParseFeed() {
  const response = UrlFetchApp.fetch(CONSTANTS.FEED_URL);
  return XmlService.parse(response.getContentText());
}

/**
 * 地震情報のエントリーをフィルタリングする
 * @param {XMLDocument} feed - パースされたXMLドキュメント
 * @returns {Array<XMLElement>} 地震情報のエントリー配列
 */
function filterEarthquakeEntries(feed) {
  const atom = XmlService.getNamespace(CONSTANTS.NAMESPACES.ATOM);
  const entries = feed.getRootElement().getChildren('entry', atom);
  
  return entries.filter(entry => {
    const content = entry.getChildText('content', atom);
    return content.includes(CONSTANTS.CONTENT_TYPE);
  });
}

/**
 * 地震情報のコンテンツを処理する
 * @param {Array<XMLElement>} earthQuakeContents - 地震情報のエントリー配列
 * @returns {Array<Array<string>>} 処理された地震情報の配列
 */
function processEarthquakeContents(earthQuakeContents) {
  const atom = XmlService.getNamespace(CONSTANTS.NAMESPACES.ATOM);
  const seismology = XmlService.getNamespace(CONSTANTS.NAMESPACES.SEISMOLOGY);

  return earthQuakeContents.map(entry => {
    try {
      const link = entry.getChild('link', atom).getAttribute('href').getValue();
      const res = UrlFetchApp.fetch(link);
      const xml = XmlService.parse(res.getContentText());
      const root = xml.getRootElement();
      const bodies = root.getChildren('Body', seismology);
      
      if (!bodies.length) {
        Logger.log('地震情報の本文が見つかりませんでした');
        return null;
      }

      const earthQuake = bodies[0].getChild('Earthquake', seismology);
      const intensity = bodies[0].getChild('Intensity', seismology);

      if (!earthQuake || !intensity) {
        Logger.log('地震情報または震度情報が見つかりませんでした');
        return null;
      }

      const originTime = earthQuake.getChildText('OriginTime', seismology);
      const areaName = earthQuake.getChild('Hypocenter', seismology)
        .getChild('Area', seismology)
        .getChildText('Name', seismology);
      const maxInt = intensity.getChild('Observation', seismology)
        .getChildText('MaxInt', seismology);

      return [originTime, areaName, maxInt];
    } catch (error) {
      Logger.log(`地震情報の処理中にエラーが発生しました: ${error.message}`);
      return null;
    }
  }).filter(Boolean);
}

/**
 * 地震情報をスプレッドシートに保存する
 * @param {Array<Array<string>>} earthquakes - 保存する地震情報の配列
 */
function saveToSpreadsheet(earthquakes) {
  if (!earthquakes?.length) {
    Logger.log('保存する地震情報がありません');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  earthquakes.forEach(quake => {
    sheet.appendRow(quake);
  });
  Logger.log(`${earthquakes.length}件の地震情報を保存しました`);
}