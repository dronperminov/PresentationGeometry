# PresentationGeometry

## Описание
Этот инструмент позволяет добавлять базовые геометрические фигуры в слайд PowerPoint (pptx). Поддерживаются линии, эллипсы, прямоугольники (включая закругленные углы) и полигоны. Работа осуществляется напрямую с XML-структурой файла pptx, что позволяет эффективно модифицировать содержимое презентации.

## Возможности
- Добавление линий, эллипсов, прямоугольников и полигонов.
- Поддержка параметров заливки, обводки, толщины линий и прозрачности.
- Возможность группировки фигур одного типа.

## Установка
Для работы инструмента не требуются внешние зависимости.

## Пример использования
```python
import tempfile
from presentation import Presentation

with tempfile.TemporaryDirectory() as work_dir:
    presentation = Presentation(presentation_path="empty.pptx", work_path=work_dir)

    presentation.add_ellipse({"x": 20, "y": 2, "d": 4, "fill": "#7699d4"})
    presentation.add_rectangle({"x": 18, "y": 8, "w": 4, "h": 8.5, "radius": 0.25, "fill": "#dd7373", "stroke": "#222", "thickness": 3, "rotate": 30})
    presentation.add_line({"x1": 1, "y1": 1, "x2": 10, "y2": 1, "thickness": 2, "stroke": "#7699d4"})

    presentation.save("result.pptx")
```

## Параметры фигур
### Линия
- `x1`, `y1` – начальная координата;
- `x2`, `y2` – конечная координата;
- `stroke` – цвет линии;
- `stroke-opacity` – непрозрачность обводки;
- `thickness` – толщина обводки.

#### Пример:

```python
{
    "x1": 1,
    "y1": 1,
    "x2": 10,
    "y2": 1,
    "stroke": "#7699d4",
    "stroke-opacity": 0.8,
    "thickness": 2
}
```

### Эллипс
- `x`, `y` – координата левого верхнего угла;
- `dx` – горизонтальный диаметр;
- `dy` – вертикальный диаметр;
- `d` – диаметр (для окружности);
- `rotate` – угол поворота (в градусах);
- `fill` – цвет заливки;
- `fill-opacity` – непрозрачность заливки;
- `stroke` – цвет обводки;
- `stroke-opacity` – непрозрачность обводки;
- `thickness` – толщина обводки.

#### Пример

```python
{
    "x": 20, "y": 2,
    "dx": 4, "dy": 6,
    "rotate": 15,
    "fill": "#7699d4",
    "fill-opacity": 0.6,
    "stroke": "#222",
    "stroke-opacity": 0.8,
    "thickness": 2
}
```

### Прямоугольник
- `x`, `y` – координата левого верхнего угла;
- `w`, `h` – ширина и высота;
- `radius` – процент скругления углов (от 0 до 1);
- `rotate` – угол поворота (в градусах);
- `fill` – цвет заливки;
- `fill-opacity` – непрозрачность заливки;
- `stroke` – цвет обводки;
- `stroke-opacity` – непрозрачность обводки;
- `thickness` – толщина обводки.

#### Пример:

```python
{
    "x": 18,
    "y": 8,
    "w": 4,
    "h": 8.5,
    "radius": 0.25,
    "rotate": 30,
    "fill": "#dd7373",
    "fill-opacity": 0.7,
    "stroke": "#222",
    "stroke-opacity": 0.9,
    "thickness": 3
}
```

### Полигон
- `points` – список координат точек (словарь вида `{"x": float, "y": float}`);
- `rotate` – угол поворота (в градусах);
- `fill` – цвет заливки;
- `fill-opacity` – непрозрачность заливки;
- `stroke` – цвет обводки;
- `stroke-opacity` – непрозрачность обводки;
- `thickness` – толщина обводки.

#### Пример:

```python
{
    "points": [
        {"x": 1, "y": 1},
        {"x": 4, "y": 5},
        {"x": 7, "y": 2}
    ],
    "rotate": 10,
    "fill": "#ffcc00",
    "fill-opacity": 0.5,
    "stroke": "#333",
    "stroke-opacity": 0.8,
    "thickness": 2
}
```

## Методы добавления фигур

### Добавление одной фигуры
- `add_line(line: dict)` – добавление линии;
- `add_ellipse(ellipse: dict)` – добавление эллипса;
- `add_rectangle(rectangle: dict)` – добавление прямоугольника;
- `add_polygon(polygon: dict)` – добавление полигона.

### Добавление нескольких фигур (групповое добавление)
- `add_lines(lines: List[dict])` – добавление линий;
- `add_ellipses(ellipses: List[dict])` – добавление эллипсов;
- `add_rectangles(rectangles: List[dict])` – добавление прямоугольников;
- `add_polygons(polygons: List[dict])` – добавление полигонов.

## Как это работает
1. Используется заранее подготовленный пустой файл `empty.pptx`.
2. Презентация распаковывается во временную директорию (pptx – это всего-лишь zip-архив).
3. Внесение изменений происходит в файле `ppt/slides/slide1.xml` (работа ведётся только с первым слайдом).
4. После внесения изменений всё запаковывается обратно в pptx-файл.

## Примеры

Примеры использования можно найти в папке `examples`

### Пример 1. Базовые фигуры

Код примера находится в файле [examples/basic.py](examples/basic.py)

![Пример слайда](examples/basic.png)

Скачать пример презентации: [examples/basic.pptx](examples/basic.pptx)


### Пример 2. Визуализация распределений

Код примера находится в файле [examples/scatter.py](examples/scatter.py)

![Пример слайда](examples/scatter.png)

Скачать пример презентации: [examples/scatter.pptx](examples/scatter.pptx)


### Пример 3. Гистограммы

Код примера находится в файле [examples/histogram.py](examples/histogram.py)

![Пример слайда](examples/histogram.png)

Скачать пример презентации: [examples/histogram.pptx](examples/histogram.pptx)


## Лицензия
MIT License
