/*
  점포/사무실 좌표 DB
  - lon: 경도(X), lat: 위도(Y)
  - TMAP 자동차 경로 API는 startX/endX=경도, startY/endY=위도 기준으로 사용합니다.
  - 대구 사무실 좌표는 현재 '코스트코 대구점 인근' 임시값입니다.
    실제 사무실 위치로 바꾸면 출발/복귀 거리 정확도가 올라갑니다.
*/
window.PLACE_DB = [
  {
    name: '대구 사무실',
    aliases: ['대구사무실', '대구 사무실', '대구오피스', '대구 오피스'],
    address: '대구광역시 북구 검단로 97 인근',
    shortAddress: '대구 북구 검단로 인근',
    lon: 128.6138,
    lat: 35.9063
  },
  {
    name: '코스트코 대구점',
    aliases: ['코스트코 대구점', '대구점', 'Costco 대구점'],
    address: '대구광역시 북구 검단로 97',
    shortAddress: '대구 북구 검단로 97',
    lon: 128.6130,
    lat: 35.9068
  },
  {
    name: '코스트코 혁신점',
    aliases: ['코스트코 혁신점', '코스트코 대구혁신점', '대구혁신점', '혁신점', 'Costco 혁신점'],
    address: '대구광역시 동구 첨단로 10',
    shortAddress: '대구 동구 첨단로 10',
    lon: 128.7272,
    lat: 35.8791
  },
  {
    name: '코스트코 대전점',
    aliases: ['코스트코 대전점', '대전점', 'Costco 대전점'],
    address: '대전광역시 중구 오류로 41',
    shortAddress: '대전 중구 오류로 41',
    lon: 127.4079,
    lat: 36.3228
  },
  {
    name: '코스트코 세종점',
    aliases: ['코스트코 세종점', '세종점', 'Costco 세종점'],
    address: '세종특별자치시 종합운동장1로 14',
    shortAddress: '세종 종합운동장1로 14',
    lon: 127.2587,
    lat: 36.5144
  },
  {
    name: '코스트코 울산점',
    aliases: ['코스트코 울산점', '울산점'],
    address: '울산광역시 북구 진장유통로 78-12',
    shortAddress: '울산 북구 진장유통로 78-12',
    lon: 129.3541,
    lat: 35.5734
  },
  {
    name: '코스트코 부산점',
    aliases: ['코스트코 부산점', '부산점'],
    address: '부산광역시 수영구 구락로 137',
    shortAddress: '부산 수영구 구락로 137',
    lon: 129.1157,
    lat: 35.1771
  },
  {
    name: '코스트코 김해점',
    aliases: ['코스트코 김해점', '김해점'],
    address: '경상남도 김해시 주촌면 선천남로 16',
    shortAddress: '김해 주촌면 선천남로 16',
    lon: 128.8297,
    lat: 35.2280
  },
  {
    name: '양산 사무실',
    aliases: ['양산사무실', '양산 사무실'],
    address: '경상남도 양산시',
    shortAddress: '경남 양산시',
    lon: 129.0370,
    lat: 35.3350
  }
];
