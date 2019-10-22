# DoExcel

#### 介绍
专注于 Excel 处理操作

#### 注解
1. @DoSheet 针对单元表

| title | 单元表的标题 |
| ----- | ---------- |
| localeResource | 国际化资源文件地址 |

2. @Column 针对实体类属性 也对应 Excel 列数据

| name        | 字段名称                   |
| ----------- | -------------------------- |
| columnEnums | 配置字段值需要转换输出注解 |

3. @ColumnEnum 当作 @Column 注解的属性 

| code        | 需要转换的字符串           |
| ----------- | -------------------------- |
| value       | 转换之后的字符串            |


#### 操作方法

#### 已实现功能
1. 注解配置，导出 Excel 文件

2. 支持国际化（单元表标题，列标题以及某个字段）

    只需要在实体类上加上 @DoSheet 注解，同时配置属性 localeResource
    然后国际化资源文件对应 @ColumnEnum 的 value 属性即可

3. 支持字段值转换输出（比如某个字段 status 取值 {0: 无效, 1: 有效} 最终导出数据用户希望看到的是字符串而非数字）

    正常导出数据是这样：

    | ID  | status |
    | --- | ------ |
    | 1   | 1   |

    用户其实想看到这样：

    | ID  | status |
    | --- | ------ |
    | 1   | 有效   |

    需要在实体类的字段里面这样：

```
@DoSheet(title = "终端信息表")
public class ExcelAgent {
    @Column(name = "id")
    private Long id;
    @Column(name = "status", columnEnums = {
            @ColumnEnum(code = "0", value = "无效"),
            @ColumnEnum(code = "1", value = "有效")
            })
    private Integer status;

    // 省略 get set
}

```

4. 支持自动计算列宽度，根据值长度自适应

5. 支持时间格式化输出


#### 待实现功能
1. 支持异步操作, 异步操作可随时取得进度描述
2. 可以预先制定使用内存大小, 防止 OOM
3. 还没想好