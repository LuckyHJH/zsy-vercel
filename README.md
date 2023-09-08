## php部署vercel总结

1. 根目录新增文件vercel.json，关键内容如下：
```
{
  "functions": {
    "api/*.php": {
      "runtime": "vercel-php@0.6.0"
    }
  }
}
```

2. 新建api/index.php，内容如下：
```
<?php
phpinfo();
```

如果是框架的，可以这样：
```
<?php
require __DIR__ . '/../public/index.php';
```

3. 访问“域名/api”即可。
