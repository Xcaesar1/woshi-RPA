export function readPageData<T>(): T {
  const element = document.getElementById("page-data");
  if (!element?.textContent) {
    throw new Error("页面初始化数据缺失");
  }
  return JSON.parse(element.textContent) as T;
}
