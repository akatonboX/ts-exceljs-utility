module.exports = {
  webpack: {
    configure: (webpackConfig, { env, paths }) => {
      //■.xlsxファイルをassetとして取り扱う
      webpackConfig.module.rules.push({
        test: /\.xlsx$/,
        type: 'asset',
        parser: {
          dataUrlCondition: {
            maxSize: 0,
          },
        },
      });
      return webpackConfig;
    },
  },
};
