import { UniverPlugin } from "@univerjs/webpack-plugin";

export default function () {
  return {
    name: "unvier-plugin",
    configureWebpack() {
      return {
        plugins: [new UniverPlugin()],
      };
    },
  };
}
